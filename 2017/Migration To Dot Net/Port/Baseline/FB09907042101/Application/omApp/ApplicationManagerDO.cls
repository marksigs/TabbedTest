VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "ApplicationManagerDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      ApplicationManagerDO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling
'
'Dependencies:  ApplicationManagerBO, ApplicationManagerTxBO
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'
'Prog  Date      Description
'SR    12/01/00  Restructured
'SR    16/03/00  SYS0265 - Delete the verification and customerrelationship records
'                before deleting customerrole records.
'JLD   02/08/00  SYS1379 - Addition of FindLegacyBusinessForCustomer
'LD    04/09/00  Ammendments for NT4/W2000 dual-compile
'LD    07/11/00  Explicity close recordsets
'LD    07/11/00  Explicity destroy command objects
'AS    15/11/00  CORE000002: Added FindApplicationsList method
'AS    15/11/00  CORE000015: Extended FindApplicationsList method
'PSC   12/01/01  SYS1749 Amend FindOmigaBusinessForCustomer to use change node name rather
'                than replace and to change the node names here rather than in the BO
'ASt   06/03/01  CC012 FindApplicationListTotals compatible with SQL Server.
'APS   11/03/01  SYS2030 Added ApplicationPriority' return parameter from FindApplicationList call
'LD    11/06/01  SYS2367 SQL Server Port - Length must be specified in calls to CreateParameter
'DM    14/06/01  SYS2375 Changed the names for the SQL Server stored procedures.
'LD    22/06/01  SYS2368 AQL Server Port - IApplicationManagerDO_FindApplicationList changed
'DM    18/07/01  SYS2493 Fixed wrong number of parameters to the Oracle FindApplicationList stored proc.
'AD    03/10/01  SYS2718 Removed MoveFirst from the ForwardOnly Cursor.
'DRC   23/04/02  SYS3017 Must delete the customerrelationship when updating roles
'MEVA  29/04/02  SYS1452 Do not report missing locks
'------------------------------------------------------------------------------------------

Option Explicit

Implements IApplicationManagerDO
Implements ObjectControl
 
'=============================================
'Constant Declaration Section
'=============================================

Private m_objContext As ObjectContext
Private m_objErrAssist As ErrAssist
Private m_objIDOAssist As IDOAssist
Private m_objXmlAssist As XMLAssist

Private Const cstrTABLE_NAME = "APPLICATION"
'Private Const cstrTABLE_APPLLOCK = "APPLICATIONLOCK"
'Private Const cstrTABLE_CUSTLOCKAPPLLOCK = "CUSTOMERLOCKAPPLICATIONLOCK"
Private Const cstrTABLE_CUSTROLE = "CUSTOMERROLE"
Private Const cintFINDAPPLICATIONLISTPAGE = 10


'------------------------------------------------------------------------------------------
'BMIDS Specific History:
'
' Prog  Date        Description
' MO    01/10/02    BMIDS00502 - Made change to the application search, FindApplicationList
'                   to cope with the new BM Introducer system.  Remove 'IntermediaryGUID'
'                   and replaced this with 'IntroducerID'
' SR    20/11/02    BMIDS00500 - added new methods 'LockCustomersForApplication',
'                   and 'UnlockApplicationAndCustomers'
' GHun  03/12/02    BM0117 - Applied Core AQR SYS5224: Existing locks should raise an error in LockCustomersForApplication
' GHun  11/12/02    BM0035 - CC011 Customer business search in application enquiry
' GHun  15/04/2003  BM0525 - Amend FindApplicationList to only make one database request
' GHun  05/06/2003  BM0523 - Split FindApplicationList to improve performance
' LDM   04/06/2003  BM0202 - show user and unit of who has locked the app
' GHun  16/06/2003  BM0523 - Amend FindApplicationList to use correct node name when searching by Approval Date
' HMA   23/07/2003  BMIDS607 - Amend FindApplicationList to convert xml to required format.
' HMA   21/07/2004  BMIDS758 - Add 'CreateRemovedToECustomer' function.
' HMA   13/09/2004  BMIDS864 - Remove ApplicationNumber and ApplicationFactFindNumber from key on Verification table
'------------------------------------------------------------------------------------------
'MARS Specific History:
'
' HMA   20/09/2005  MAR46   - Add Product Switch and ToE applications to FindApplicationList
' PJO   30/11/2005  MAR741  - Fix USP calls with currect User ID length
' PSC   25/03/2006  MAR1486 - Amend IApplicationManagerDO_LockCustomersForApplication() for
'                             performance
' GHun  13/06/2006  MAR1870 - Changed LockCustomersForApplication and UnlockApplicationAndCustomers
'------------------------------------------------------------------------------------------
' Epsom Specific History:
'
' GHun  20/12/2006  EP2_56  Changed CreateRemovedToECustomer to set SequenceNumber
'------------------------------------------------------------------------------------------
Private Sub IApplicationManagerDO_CreateCustomerLockApplicationLock( _
                                          ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   create an instance of the persistant data associated with the table
'   'CustomerLockApplicationLock' for each set of data in the request
' pass:
'   vxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo CreateVbErr

    Const strFunctionName As String = "IApplicationManagerDO_CreateCustomerLockApplicationLock"
   
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
    
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadCustomerLockApplicationLockData()
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

Private Sub IApplicationManagerDO_CreateCustomerRole( _
                        ByVal vxmlTableElement As IXMLDOMElement)
' header -----------------------------------------------------------------------
' description:
'   create an instance of the persistant data associated with this data object
'   for each set of data in the request
' pass:
'   vxmlTableElement  xml element containing the request
' return:                n/a
'--------------------------------------------------------------------------------
On Error GoTo CreateVbErr

    Const strFunctionName As String = "IApplicationManagerDO_CreateCustomerRole"
   
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

Private Sub IApplicationManagerDO_CreateLock( _
                    ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   create an instance of the persistant data associated with the table 'ApplicationLock'
'   for each set of data in the request
' pass:
'   vxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo CreateLockVbErr

    Const strFunctionName As String = "IApplicationManagerDO_CreateLock"
   
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
    
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadApplicationLockData()
    m_objIDOAssist.Create vxmlTableElement, xmlClassDefDoc
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

    Exit Sub

CreateLockVbErr:
    
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

Private Sub IApplicationManagerDO_DeleteAllCustomerLockApplicationLock( _
                ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   Delete all instances of the persistant data associated with the table
'   'CustomerLockApplicationLock'
' pass:
'   vxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo DeleteVbErr

    Const strFunctionName As String = "IApplicationManagerDO_DeleteAllCustomerLockApplicationLock"
    
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
                
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadCustomerLockApplicationLockData()
    m_objIDOAssist.DeleteAll vxmlTableElement, xmlClassDefDoc
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Exit Sub

DeleteVbErr:

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

Private Sub IApplicationManagerDO_DeleteCustomerLockApplicationLock(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   Delete a single instance of the persistant data associated with the table
'   'CustomerLockApplicationLock'
' pass:
'   vxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo DeleteVbErr

    Const strFunctionName As String = "IApplicationManagerDO_DeleteCustomerLockApplicationLock"
    
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
                
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadCustomerLockApplicationLockData()
    m_objIDOAssist.Delete vxmlTableElement, xmlClassDefDoc
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Exit Sub

DeleteVbErr:

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

Private Sub IApplicationManagerDO_DeleteCustomerRole( _
                ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   Delete a single instance of the persistant data associated with the table
'   'CustomerRole'
' pass:
'   vxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo DeleteVbErr

    Const strFunctionName As String = "IApplicationManagerDO_DeleteCustomerRole"
    
            
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
    Dim objIADOAssist As IADOAssist
    Dim objSqlAssist As SQLAssist
                
      ' delete the  related records from the table 'CUSTOMERRELATIONSHIP', if any
    Dim strCondition As String, strApplicationNumber As String, strAFFNumber As String
    Dim strCustomerNumber As String, strCustomrVersionNumber As String
    Dim strSQL As String
    
    Set objSqlAssist = New SQLAssist
    
    If m_objContext Is Nothing Then
        Set objIADOAssist = New ADOAssist
    Else
        Set objIADOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    End If
    
    strApplicationNumber = vxmlTableElement.selectSingleNode(".//APPLICATIONNUMBER").Text
    strAFFNumber = vxmlTableElement.selectSingleNode(".//APPLICATIONFACTFINDNUMBER").Text
    strCustomerNumber = vxmlTableElement.selectSingleNode(".//CUSTOMERNUMBER").Text
    strCustomrVersionNumber = vxmlTableElement.selectSingleNode(".//CUSTOMERVERSIONNUMBER").Text
    
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
                 
    ' BMIDS864  Remove Application Number and ApplicationFactFindNumber as they can be null
    ' delete the verification records from the table 'VERIFICATION', if any
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

Private Sub IApplicationManagerDO_DeleteLock(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   Delete a single instance of the persistant data associated with the table
'   'ApplicationLock'
' pass:
'   vxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo DeleteLockVbErr

    Const strFunctionName As String = "IApplicationManagerDO_DeleteLock"
    
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
                
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadApplicationLockData()
    m_objIDOAssist.Delete vxmlTableElement, xmlClassDefDoc
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Exit Sub

DeleteLockVbErr:

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

Private Function IApplicationManagerDO_FindCustomerLockApplicationLockList( _
                        ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for all instances of the persistant data associated with the table
'   'CustomerLockApplicationLock' for the values supplied
' pass:
'   vxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo FindListVbErr

    Const strFunctionName As String = "IApplicationManagerDO_FindCustomerLockApplicationLockList"
    
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
         
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadCustomerLockApplicationLockData()
    Set IApplicationManagerDO_FindCustomerLockApplicationLockList = m_objIDOAssist.FindList(vxmlTableElement, xmlClassDefDoc)

    
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
    
    If Err.Description <> "Record not found" Then
        '
        '   re-raise error for business object to interpret as appropriate
        '
        Err.Raise Err.Number, Err.Source, Err.Description
    End If

End Function

Private Function IApplicationManagerDO_FindCustomerRoleList( _
                    ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ------------------------------------------------------------------------------
' description:
'   Get the data for all instances of the persistant data associated with the table
'   'CustomerRole' for the values supplied
' pass:
'   vxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'--------------------------------------------------------------------------------------
On Error GoTo FindListVbErr

    Const strFunctionName As String = "IApplicationManagerDO_FindCustomerRoleList"
    
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
    Dim strOrderBy As String
         
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadCustomerRoleData()
    strOrderBy = " CUSTOMERROLETYPE ASC, CUSTOMERORDER ASC "
    Set IApplicationManagerDO_FindCustomerRoleList = _
            m_objIDOAssist.FindList(vxmlTableElement, xmlClassDefDoc, strOrderBy)

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

Private Function IApplicationManagerDO_FindLegacyBusinessForCustomer(ByVal vxmlRequest As IXMLDOMNode) As IXMLDOMElement
' header ----------------------------------------------------------------------------------
' description: Get omiga business for a particular customer and customer version
'              combination
' pass:
'   vstrXMLRequest  string containing the xml Request containing data to which identifies
'                   the instance(s) of the persistant data to be retrieved
'                   Request format:
'                   <CUSTOMER>number</CUSTOMER>
' return:
'   FindList        XML node representation of  data retrieved
'                   Response format:
'                        <FINDLEGACYBUSINESSFORCUSTOMERLIST>
'                          <FINDLEGACYBUSINESSFORCUSTOMER>
'                              <CUSTOMERNUMBER></CUSTOMERNUMBER>
'                              <CUSTOMERVERSIONNUMBER></CUSTOMERVERSIONNUMBER>
'                              <APPLICATIONNUMBER></APPLICATIONNUMBER>
'                              <TYPEOFAPPLICATION></TYPEOFAPPLICATION>
'                              <BUSINESSTYPEINDICATOR></BUSINESSTYPEINDICATOR>
'                              <APPLICATIONFACTFINDNUMBER></APPLICATIONFACTFINDNUMBER>
'                              <DATECREATED></DATECREATED>
'                              <PACKAGENUMBER></PACKAGENUMBER>
'                              <CORRESPONDENCESALUTATION></CORRESPONDENCESALUTATION>
'                              <SURNAME></SURNAME>
'                              <FIRSTFORENAME></FIRSTFORENAME>
'                              <DATEOFBIRTH></DATEOFBIRTH>
'                              <AMOUNT></AMOUNT>
'                              <STATUS></STATUS>
'                              <TOTALLOANAMOUNT></TOTALLOANAMOUNT>
'                              <AMOUNTREQUESTED></AMOUNTREQUESTED>
'                          </FINDLEGACYBUSINESSFORCUSTOMER>
'                        </FINDLEGACYBUSINESSFORCUSTOMERLIST>
' Raise Errors:
'   if record not found, raise omiga4RecordNotFound
'------------------------------------------------------------------------------------------
On Error GoTo FindLegacyBusinessForCustomerVbErr

    Const strFunctionName As String = "IApplicationManagerDO_FindLegacyBusinessForCustomer"

    Dim objErrAssist As New ErrAssist
    
    Dim objIDOAssist As IDOAssist
    Dim objXmlAssist As New XMLAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim xmlDoc
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
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
    
    Set objIApplicationClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIApplicationClassDef.LoadFindLegacyBusinessForCustomerData()
    
    ' Find the additional condition due to ApplicationSatge
    Dim strFromApplStage As String, strToApplStage As String, strAdditionalCondition As String
    strFromApplStage = objXmlAssist.GetElementText(vxmlRequest, "//APPLICATIONSTAGEFROM")
    strToApplStage = objXmlAssist.GetElementText(vxmlRequest, "//APPLICATIONSTAGETO")

    strAdditionalCondition = IIf(strFromApplStage <> "", " STAGENUMBER >= " & strFromApplStage, "")
    
    If Len(strToApplStage) <> 0 Then
        strAdditionalCondition = _
                strAdditionalCondition & IIf(strAdditionalCondition <> "", " AND ", "")
        strAdditionalCondition = _
                strAdditionalCondition & " STAGENUMBER <= " & strToApplStage
    End If
        
    Dim strOtherSystemCustomerNumber As String
    strOtherSystemCustomerNumber = objXmlAssist.GetMandatoryElementText(vxmlRequest, "//OTHERSYSTEMCUSTOMERNUMBER")
    strAdditionalCondition = strAdditionalCondition & IIf(strAdditionalCondition <> "", " AND ", "")
    strAdditionalCondition = strAdditionalCondition & " OTHERSYSTEMCUSTOMERNUMBER = '" & strOtherSystemCustomerNumber & "'"
    
    strOrderBy = " DATECREATED DESC, APPLICATIONNUMBER ASC "

    Set xmlNode = _
            objIDOAssist.FindListMultiple(vxmlRequest, xmlClassDefDoc, _
                                          strOrderBy, , strAdditionalCondition, " AND ")
    
    Set xmlNodeList = xmlNode.selectNodes(".//" & "FINDLEGACYBUSINESSFORCUSTOMER")
    
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
    
    Set IApplicationManagerDO_FindLegacyBusinessForCustomer = xmlNode

    Set xmlDoc = Nothing
    Set objXmlAssist = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlNode = Nothing
    Set xmlElement = Nothing
    Set xmlBusinessNode = Nothing
    Set objIApplicationClassDef = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

FindLegacyBusinessForCustomerVbErr:
    
    Set xmlDoc = Nothing
    Set objXmlAssist = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
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

Private Function IApplicationManagerDO_FindOmigaBusinessForCustomer(ByVal vstrXMLRequest As String) As IXMLDOMElement
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

    Const strFunctionName As String = "IApplicationManagerDO_FindOmigaBusinessForCustomer"

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
        
    Set xmlDoc = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    m_objXmlAssist.ChangeNodeName xmlDoc.documentElement, "CUSTOMER", "FINDOMIGABUSINESSFORCUSTOMER"
     
    Set objIApplicationClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIApplicationClassDef.LoadFindOmigaBusinessForCustomerData()
    
    ' Find the additional condition due to ApplicationSatge
    Dim strFromApplStage As String, strToApplStage As String, strAdditionalCondition As String
    strFromApplStage = objXmlAssist.GetElementText(xmlDoc.documentElement, "//APPLICATIONSTAGEFROM")
    strToApplStage = objXmlAssist.GetElementText(xmlDoc.documentElement, "//APPLICATIONSTAGETO")

    
    strAdditionalCondition = IIf(strFromApplStage <> "", " STAGENUMBER >= " & strFromApplStage, "")
    
    If Len(strToApplStage) <> 0 Then
        strAdditionalCondition = _
                strAdditionalCondition & IIf(strAdditionalCondition <> "", " AND ", "")
        strAdditionalCondition = _
                strAdditionalCondition & " STAGENUMBER <= " & strToApplStage
    End If
        
    strOrderBy = " DATECREATED DESC, APPLICATIONNUMBER ASC "

    Set xmlNode = _
            objIDOAssist.FindListMultiple(xmlDoc.documentElement, xmlClassDefDoc, _
                                          strOrderBy, , strAdditionalCondition, " AND ")
    
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
    
    m_objXmlAssist.ChangeNodeName xmlNode, "FINDOMIGABUSINESSFORCUSTOMER", "OMIGABUSINESS"
    m_objXmlAssist.ChangeNodeName xmlNode, "FINDOMIGABUSINESSFORCUSTOMERLIST", "OMIGABUSINESSLIST"
    
    Set IApplicationManagerDO_FindOmigaBusinessForCustomer = xmlNode

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

Private Function IApplicationManagerDO_GetCustomerRoleData( _
                    ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with the
'   table 'CustomerRole'
' pass:
'   vxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo GetDataVbErr

    Const strFunctionName As String = "IApplicationManagerDO_GetCustomerRoleData"
      
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
    
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadCustomerRoleData()
    Set IApplicationManagerDO_GetCustomerRoleData = _
                m_objIDOAssist.GetData(vxmlTableElement, xmlClassDefDoc)
    
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

Private Function IApplicationManagerDO_GetLockData( _
                ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   the table 'ApplicationLock'
' pass:
'   vxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo GetLockDataVbErr

    Const strFunctionName As String = "IApplicationManagerDO_GetLockData"
      
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
    
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadApplicationLockData()
    Set IApplicationManagerDO_GetLockData = _
                        m_objIDOAssist.GetData(vxmlTableElement, xmlClassDefDoc)
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    
'    If Not m_objContext Is Nothing Then
'        m_objContext.SetComplete
'    End If
    
    Exit Function

GetLockDataVbErr:
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    
    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
'    If Not m_objContext Is Nothing Then
'        If m_objErrAssist.IsSystemError Then
'            m_objContext.SetAbort
'        Else
'            m_objContext.SetComplete
'        End If
'    End If
        
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Sub IApplicationManagerDO_LockCustomersForApplication(ByVal vxmlRequest As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   Places an application lock and related customer locks into the database via a stored-
'   procedure.
'------------------------------------------------------------------------------------------
On Error GoTo LockCustomersForApplicationErr

    Const strFunctionName As String = "IApplicationManagerDO_LockCustomersForApplication"
       
    Dim sUserID As String
    Dim sUnitID As String
    Dim sMachineID As String
    Dim sApplicationNumber As String
    Dim oCommand        As ADODB.Command
    Dim adoConnection   As ADODB.Connection ' PSC 25/03/2006 MAR1486
    'MAR1870 GHun
    Dim adoStream       As ADODB.Stream
    Dim xmlDoc          As FreeThreadedDOMDocument40
    'MAR1870 End
    
    'Use XML assist and obtain the parameters to pass into the command object.
    sUserID = xmlGetMandatoryAttributeText(vxmlRequest, "USERID")
    sUnitID = xmlGetMandatoryAttributeText(vxmlRequest, "UNITID")
    sMachineID = xmlGetMandatoryAttributeText(vxmlRequest, "MACHINEID")
    sApplicationNumber = xmlGetMandatoryNodeText(vxmlRequest, "//APPLICATIONNUMBER")

    ' PSC 25/03/2006 MAR1486 - Start
    Set adoConnection = New ADODB.Connection
    With adoConnection
        .ConnectionString = adoGetDbConnectString
        .open
    End With
    ' PSC 25/03/2006 MAR1486 - End
    
    'MAR1870 GHun
    Set adoStream = New ADODB.Stream
    adoStream.open
    'MAR1870 End
    
    'Create a command object to use.
    Set oCommand = New ADODB.Command
    
    ' PSC 25/03/2006 MAR1486
    With oCommand
        Set .ActiveConnection = adoConnection

        .CommandType = adCmdStoredProc
        .CommandText = "USP_APPLICATION_LOCKAPP_WITHFAILLOCKINFO" 'BM0202 21/05/03 LDM
        ' PSC 25/03/2006 MAR1486
    
        'PJO 30/11/2005 MAR741 - Fix lengths for user IDs
        .Parameters.Append .CreateParameter("@USERID", adVarChar, adParamInput, Len(sUserID), sUserID)
        .Parameters.Append .CreateParameter("@UNITID", adVarChar, adParamInput, Len(sUnitID), sUnitID)
        .Parameters.Append .CreateParameter("@MACHINEID", adVarChar, adParamInput, Len(sMachineID), sMachineID)
        .Parameters.Append .CreateParameter("@APPLICATIONNUMBER", adVarChar, adParamInput, Len(sApplicationNumber), sApplicationNumber)
    
        'BM0117 - Existing locks should raise an error.
        .Parameters.Append .CreateParameter("@APPLICATIONALREADYLOCKED", adVarChar, adParamOutput, 10)
        .Parameters.Append .CreateParameter("@CUSTOMERALREADYLOCKED", adBoolean, adParamOutput)
        'BM0117 End

        'BM0202 21/05/03 LDM parameters to recieve all which unit the user who has locked the app belongs to &
        'customers which are locked and who by
        .Parameters.Append .CreateParameter("@APPLICATIONUNITIDLOCKED", adVarChar, adParamOutput, 10)
    
        'MAR1870 GHun This XML is now returned as a stream
        '.Parameters.Append .CreateParameter("@CUSTOMERXML", adVarChar, adParamOutput, 1500)

        .Properties("Output Stream") = adoStream
        .Properties("XML Root") = "CUSTOMERLIST"

        'Execute the command/stored proc and lock the application and customers.
        .Execute , , adExecuteStream
    
        ' PSC 25/03/2006 MAR1486 - Start
        'Close the database connection
        Set .ActiveConnection = Nothing
    End With
    
    adoConnection.Close
    ' PSC 25/03/2006 MAR1486 - End

    'BM0117 Check the output parameters to determine if any locks previously existed.
    If IsNull(oCommand.Parameters("@APPLICATIONALREADYLOCKED").Value) = False Then
        'The application is already locked by %s...
        'BM0202 21/05/03 LDM show user and unit of who has locked the app
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 7014, , _
                                    oCommand.Parameters("@APPLICATIONALREADYLOCKED").Value, _
                                    oCommand.Parameters("@APPLICATIONUNITIDLOCKED").Value
    ElseIf IsNull(oCommand.Parameters("@CUSTOMERALREADYLOCKED").Value) = False Then
        'BM0202 21/05/03 LDM show all users who are locked
        
        'MAR1870 GHun
        Set xmlDoc = xmlCreateDOMObject
        If xmlDoc.loadXML(adoStream.ReadText) Then
        'MAR1870 End
        
            Dim xmlNode As IXMLDOMNode
            Dim sCustID As String
            Dim sCustUnitID As String
            Dim sCustUserID As String
            Dim sCustInfo As String
            
            sCustInfo = vbCrLf & vbCrLf
            
            For Each xmlNode In xmlDoc.documentElement.selectNodes("Customer")
                sCustID = xmlGetAttributeText(xmlNode, "CustID")
                sCustUnitID = xmlGetAttributeText(xmlNode, "UnitID")
                sCustUserID = xmlGetAttributeText(xmlNode, "UserID")
                sCustInfo = sCustInfo + "Customer No: " & sCustID & _
                            " by " & sCustUserID & " (" & sCustUnitID & ")" & vbCrLf
            Next
            sCustInfo = sCustInfo & vbCrLf
            m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 7013, , vbCrLf, sCustInfo
        
        Else
           'One of the customers is already locked...
           m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 190
        End If
    End If
    'BM0117 End
    
    'MAR1870 GHun
    If adoStream.State = adStateOpen Then
        adoStream.Close
    End If
    'MAR1870 End
    
    ' PSC 25/03/2006 MAR1486 - Start
    Set oCommand = Nothing
    Set xmlNode = Nothing
    Set adoConnection = Nothing
    ' PSC 25/03/2006 MAR1486 - End
    Set adoStream = Nothing 'MAR1870 GHun
    Set xmlDoc = Nothing    'MAR1870 GHun
  
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

    Exit Sub

LockCustomersForApplicationErr:
    
    ' PSC 25/03/2006 MAR1486 - Start
    If adoConnection.State = adStateOpen Then
        adoConnection.Close
    End If
    
    'MAR1870 GHun
    If Not adoStream Is Nothing Then
        If adoStream.State = adStateOpen Then
            adoStream.Close
        End If
    End If
    
    Set adoStream = Nothing
    Set xmlDoc = Nothing
    'MAR1870 End
    
    Set oCommand = Nothing
    Set xmlNode = Nothing
    Set adoConnection = Nothing
    ' PSC 25/03/2006 MAR1486 - End
    
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

Private Sub IApplicationManagerDO_UnlockApplicationAndCustomers(ByVal vxmlRequest As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   Removes all the locks on the given application and any of its customers.
'------------------------------------------------------------------------------------------
On Error GoTo UnlockApplicationAndCustomersErr

    Const strFunctionName As String = "IApplicationManagerDO_UnlockApplicationAndCustomers"
       
    Dim sApplicationNumber As String
    Dim oCommand As ADODB.Command
    Dim adoConn As ADODB.Connection 'MAR1870
   
    'Use XML assist and obtain the parameters to pass into the command object.
    sApplicationNumber = xmlGetMandatoryNodeText(vxmlRequest, "//APPLICATIONNUMBER")

    'MAR1870 GHun Explicitly open connection
    Set adoConn = New ADODB.Connection
    With adoConn
        .ConnectionString = adoGetDbConnectString
        .open
    End With
    'MAR1870 End

    'Create a command object to use.
    Set oCommand = New ADODB.Command
    With oCommand
        .CommandType = adCmdStoredProc
        .CommandText = "USP_APPLICATION_UNLOCKCUSTOMERSFORAPPLICATION"
        Set .ActiveConnection = adoConn
        .Parameters.Append .CreateParameter("@APPLICATIONNUMBER", adVarChar, adParamInput, Len(sApplicationNumber), sApplicationNumber)
       
        'Execute the command/stored proc and unlock the application and customers.
        .Execute , , adExecuteNoRecords
        Set .ActiveConnection = Nothing 'MAR1870 GHun
    End With
        
    adoConn.Close   'MAR1870 GHun
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

    Set adoConn = Nothing   'MAR1870 GHun
    Set oCommand = Nothing  'MAR1870 GHun

    Exit Sub

UnlockApplicationAndCustomersErr:
    'MAR1870 GHun
    If adoConn.State = adStateOpen Then
        adoConn.Close
    End If
    Set adoConn = Nothing
    Set oCommand = Nothing
    'MAR1870 End
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

Private Sub IApplicationManagerDO_UpdateCustomerRole(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   Update a single instance of the persistant data associated with the table
'   'CustomerRole'
' pass:
'   vxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo UpdateVbErr

    Const strFunctionName As String = "IApplicationManagerDO_UpdateCustomerRole"
            
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
  'AQR SYS3017 start
      ' delete record from the table 'CUSTOMERRELATIONSHIP', if any
    Dim objIADOAssist As IADOAssist
    Dim objSqlAssist As SQLAssist
    Dim strCondition As String, strApplicationNumber As String, strAFFNumber As String
    Dim strCustomerNumber As String, strCustomrVersionNumber As String
    Dim strSQL As String
    
    Set objSqlAssist = New SQLAssist
    
    If m_objContext Is Nothing Then
        Set objIADOAssist = New ADOAssist
    Else
        Set objIADOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    End If
    
    strApplicationNumber = vxmlTableElement.selectSingleNode(".//APPLICATIONNUMBER").Text
    strAFFNumber = vxmlTableElement.selectSingleNode(".//APPLICATIONFACTFINDNUMBER").Text
    strCustomerNumber = vxmlTableElement.selectSingleNode(".//CUSTOMERNUMBER").Text
    strCustomrVersionNumber = vxmlTableElement.selectSingleNode(".//CUSTOMERVERSIONNUMBER").Text
    
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
    
    On Error GoTo UpdateVbErr
    If m_objErrAssist.LastErr <> 0 And _
        m_objErrAssist.GetOmigaErrorNumber(m_objErrAssist.LastErr) <> oeNoRowsAffected Then
        m_objErrAssist.ReRaise
    End If
'AQR SYS3017 end
    
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

Private Function IApplicationManagerDO_FindMainApplicants( _
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

    Const strFunctionName As String = "IApplicationManagerDO_FindMainApplicants"

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

    strSQL = "select * from " & cstrTABLE_CUSTROLE & " where " & strSearchCriteria _
             & " Order By CUSTOMERORDER"

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

    ' AD 3/10/01 - Removed  the MoveFirst (ForwardOnly cursor by default). SYS2718

    While Not rstThisRecSet.EOF
        ' convert record set to xml
        Set objIDOAssist = New DOAssist
        Set objIClassDef = New ApplicationClassDef
        objIDOAssist.GetXMLFromRecordSet rstThisRecSet, objIClassDef.LoadCustomerRoleData, objXmlListNode
        
        rstThisRecSet.MoveNext
    Wend
        rstThisRecSet.Close

    Set IApplicationManagerDO_FindMainApplicants = objXmlOut.documentElement

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
    
    Set m_objXmlAssist = New XMLAssist
    
End Sub

Private Function ObjectControl_CanBePooled() As Boolean

    ObjectControl_CanBePooled = False

End Function

Private Sub ObjectControl_Deactivate()

    Set m_objContext = Nothing
    Set m_objErrAssist = Nothing
    Set m_objIDOAssist = Nothing
    Set m_objXmlAssist = Nothing

End Sub

'MO - 01/10/2002 - BMIDS00502, INWP1
Private Function IApplicationManagerDO_FindApplicationList(ByVal vobjXmlRequest As IXMLDOMNode) _
                                     As IXMLDOMElement
    
    Const cstrFunctionName As String = "IApplicationManagerDO_FindApplicationList"
    
    Dim objADOAssist                As ADOAssist
    Dim objConvertAssist            As ConvertAssist
    Dim objSqlAssist                As SQLAssist
    Dim adoConnection               As ADODB.Connection
    Dim adoCommand                  As ADODB.Command
    Dim adoParameter                As ADODB.Parameter
    Dim adoTotalRecordsParam        As ADODB.Parameter
    Dim adoStream                   As ADODB.Stream
    Dim xmlDoc                      As FreeThreadedDOMDocument40
    Dim xmlNode                     As IXMLDOMNode
    Dim xmlTypeOfAppText            As IXMLDOMNode
    Dim xmlTypeOfApp                As IXMLDOMElement
    Dim lngStartRow                 As Long
    Dim lngPageSize                 As Long
    Dim lngTotalRecords             As Long
    Dim strSQL                      As String
    Dim strApplicationNumber        As String
    Dim strAccountNumber            As String
    Dim strChannelID                As String
    Dim strDepartmentId             As String
    Dim strUserId                   As String
    Dim strUnitId                   As String
    Dim strMortgageProductCode      As String
    Dim strMortgageProductStartDate As String
    Dim strTypeOfApplication        As String
    Dim intTypeOfApplication        As Integer
    Dim strApplicationFromStage     As String
    Dim strApplicationToStage       As String
    Dim strApplicationDateFrom      As String
    Dim strApplicationDateTo        As String
    Dim strApprovalMonth            As String
    Dim strApprovalYear             As String
    Dim strIntroducerID             As String
    Dim strIncludeCancelledApps     As String
    Dim strIncludeDeclinedApps      As String
    Dim strDirectIndirectBusiness   As String
    Dim strSurname                  As String
    Dim strPostcode                 As String
    Dim blnIsWildCard               As Boolean
    
    'MAR46
    Dim strIncludeProductSwitchApps As String
    Dim strIncludeToEApps           As String
    
On Error GoTo FindApplicationListVbErr
                                    
    'BM0523 Decide which stored proc to call based on whether application number or surname are passed in
    strApplicationNumber = m_objXmlAssist.GetTagValue(vobjXmlRequest, "APPLICATIONNUMBER")
    strSurname = m_objXmlAssist.GetTagValue(vobjXmlRequest, "SURNAME")
                                    
    If Len(strApplicationNumber) > 0 Then
        Set IApplicationManagerDO_FindApplicationList = FindApplicationListByAppNum(vobjXmlRequest)
    ElseIf Len(strSurname) > 0 Then
        Set IApplicationManagerDO_FindApplicationList = FindApplicationListByCustomer(vobjXmlRequest)
    Else
    'BM0523 End
                                
        Set objADOAssist = New ADOAssist
        Set objSqlAssist = New SQLAssist
        Set adoConnection = New ADODB.Connection
        Set adoCommand = New ADODB.Command
        
        If m_objContext Is Nothing Then
            Set objConvertAssist = New DOAssist
        Else
            Set objConvertAssist = m_objContext.CreateInstance(App.Title & ".ConvertAssist")
        End If
            
        Set xmlNode = vobjXmlRequest.selectSingleNode("//REQUEST")
        
        'Determine the start record and number of records to return
        GetRecordRange xmlNode.cloneNode(False), lngStartRow, lngPageSize
                  
        strAccountNumber = m_objXmlAssist.GetTagValue(vobjXmlRequest, "ACCOUNTNUMBER")
        strDepartmentId = m_objXmlAssist.GetTagValue(vobjXmlRequest, "DEPARTMENTID")
        strUserId = m_objXmlAssist.GetTagValue(vobjXmlRequest, "USERID")
        strUnitId = m_objXmlAssist.GetTagValue(vobjXmlRequest, "UNITID")
        strMortgageProductCode = m_objXmlAssist.GetTagValue(vobjXmlRequest, "MORTGAGEPRODUCTCODE")
        strMortgageProductStartDate = m_objXmlAssist.GetTagValue(vobjXmlRequest, "MORTGAGEPRODUCTSTARTDATE")
        strTypeOfApplication = m_objXmlAssist.GetTagValue(vobjXmlRequest, "APPLICATIONTYPE")
        strApplicationFromStage = m_objXmlAssist.GetTagValue(vobjXmlRequest, "APPLICATIONSTAGEFROM")
        strApplicationToStage = m_objXmlAssist.GetTagValue(vobjXmlRequest, "APPLICATIONSTAGETO")
        strApplicationDateFrom = m_objXmlAssist.GetTagValue(vobjXmlRequest, "APPLICATIONDATEFROM")
        strApplicationDateTo = m_objXmlAssist.GetTagValue(vobjXmlRequest, "APPLICATIONDATETO")
        strApprovalYear = m_objXmlAssist.GetTagValue(vobjXmlRequest, "APPLICATIONAPPROVEDYEAR")
        'MO - 01/10/2002 - BMIDS00502, INWP1 - Start
        'strIntermediaryGUID = m_objXmlAssist.GetTagValue(vobjXmlRequest, "INTERMEDIARYGUID")
        strIntroducerID = m_objXmlAssist.GetTagValue(vobjXmlRequest, "INTRODUCERID")
        'MO - 01/10/2002 - BMIDS00502, INWP1 - End
        strIncludeCancelledApps = m_objXmlAssist.GetTagValue(vobjXmlRequest, "INCLUDECANCELLEDAPPS")
        strIncludeDeclinedApps = m_objXmlAssist.GetTagValue(vobjXmlRequest, "INCLUDEDECLINEDAPPS")
        'MAR46
        strIncludeProductSwitchApps = m_objXmlAssist.GetTagValue(vobjXmlRequest, "INCPRODUCTSWITCHAPPSCHECKED")
        strIncludeToEApps = m_objXmlAssist.GetTagValue(vobjXmlRequest, "INCTOEAPPSCHECKED")
        
        'MO - 01/10/2002 - BMIDS00502, INWP1
        strDirectIndirectBusiness = m_objXmlAssist.GetTagValue(vobjXmlRequest, "DIRECTINDIRECTBUSINESS")
        strPostcode = m_objXmlAssist.GetTagValue(vobjXmlRequest, "POSTCODE")
        strPostcode = objSqlAssist.FormatWildcardedString(strPostcode, blnIsWildCard, True)
        
        With adoConnection
            .ConnectionString = objADOAssist.GetConnStr
            .CursorLocation = adUseClient
            .open
        End With
    
    #If GENERIC_SQL Then
        Select Case objADOAssist.GetDBEngine
        Case SQLServer
            'DM        14/06/01  SYS2375
            strSQL = "USP_FINDAPPLICATIONLIST"
            adoCommand.CommandType = adCmdStoredProc
        Case Oracle
            'DM 18/07/01 SYS2493
            strSQL = "{call sp_FindApplicationList_gs.FindApplicationList(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,"
            strSQL = strSQL & "{RESULTSET " & CStr(lngPageSize + 1) & ","
            strSQL = strSQL & "APPLICATIONNUMBER,"
            strSQL = strSQL & "APPLICATIONFACTFINDNUMBER,"
            strSQL = strSQL & "TYPEOFAPPLICATION,"
            strSQL = strSQL & "APPLICATIONDATE,"
            strSQL = strSQL & "APPLICATIONSTAGENAME,"
            strSQL = strSQL & "APPLICATIONSTAGENUMBER,"
            strSQL = strSQL & "APPLICATIONSTAGESEQUENCENO,"
            strSQL = strSQL & "APPLICATIONLOCKED,"
            strSQL = strSQL & "TOTALLOANAMOUNT,"
            strSQL = strSQL & "PRODUCTNAME,"
            strSQL = strSQL & "CORRESPONDENCESALUTATION,"
            strSQL = strSQL & "APPLICATIONPRIORITYVALUE"
            strSQL = strSQL & "})}"
            adoCommand.CommandType = adCmdText
        End Select
    #Else
        strSQL = "{call sp_FindApplicationList.FindApplicationList(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,"
        strSQL = strSQL & "{RESULTSET " & CStr(lngPageSize + 1) & ","
        strSQL = strSQL & "APPLICATIONNUMBER,"
        strSQL = strSQL & "APPLICATIONFACTFINDNUMBER,"
        strSQL = strSQL & "TYPEOFAPPLICATION,"
        strSQL = strSQL & "APPLICATIONDATE,"
        strSQL = strSQL & "APPLICATIONSTAGENAME,"
        strSQL = strSQL & "APPLICATIONSTAGENUMBER,"
        strSQL = strSQL & "APPLICATIONSTAGESEQUENCENO,"
        strSQL = strSQL & "APPLICATIONLOCKED,"
        strSQL = strSQL & "TOTALLOANAMOUNT,"
        strSQL = strSQL & "PRODUCTNAME,"
        strSQL = strSQL & "CORRESPONDENCESALUTATION,"
        strSQL = strSQL & "APPLICATIONPRIORITYVALUE"
        strSQL = strSQL & "})}"
        adoCommand.CommandType = adCmdText
    #End If
    
        'BM0525 Opened later on for performance reasons
        'Set adoCommand.ActiveConnection = adoConnection
        'BM0525 End
    
        adoCommand.CommandText = strSQL
    
        ' Start Record
        Set adoParameter = adoCommand.CreateParameter(, adInteger, adParamInput, , lngStartRow)
        adoCommand.Parameters.Append adoParameter
    
        ' Maximum number of records to return
        Set adoParameter = adoCommand.CreateParameter(, adInteger, adParamInput, , lngPageSize)
        adoCommand.Parameters.Append adoParameter
        
        'BM0523 Searching by application number is now done in FindApplicationListByAppNum
        ' Application Number
        'Set adoParameter = adoCommand.CreateParameter("APPLICATIONNUMBER", adBSTR, adParamInput, Len(strApplicationNumber))
        'adoParameter.Attributes = adParamNullable
        'If Len(strApplicationNumber) > 0 Then
        '    adoParameter.Value = strApplicationNumber
        'Else
        '    adoParameter.Value = Null
        'End If
        'adoCommand.Parameters.Append adoParameter
        'BM0523 End
        
        ' Account Number
        Set adoParameter = adoCommand.CreateParameter("ACCOUNTNUMBER", adBSTR, adParamInput, Len(strAccountNumber))
        adoParameter.Attributes = adParamNullable
        If Len(strAccountNumber) > 0 Then
            adoParameter.Value = strAccountNumber
        Else
            adoParameter.Value = Null
        End If
        adoCommand.Parameters.Append adoParameter
    
        ' BM0523 ChannelID is never used in the search
        ' Channel ID
        'Set adoParameter = adoCommand.CreateParameter("CHANNELID", adBSTR, adParamInput, Len(strChannelID))
        'adoParameter.Attributes = adParamNullable
        'If Len(strChannelID) > 0 Then
        '    adoParameter.Value = strChannelID
        'Else
        '    adoParameter.Value = Null
        'End If
        'adoCommand.Parameters.Append adoParameter
        'BM0523 End
    
        ' Department ID
        Set adoParameter = adoCommand.CreateParameter("DEPARTMENTID", adBSTR, adParamInput, Len(strDepartmentId))
        adoParameter.Attributes = adParamNullable
        If Len(strDepartmentId) > 0 Then
            adoParameter.Value = strDepartmentId
        Else
            adoParameter.Value = Null
        End If
        adoCommand.Parameters.Append adoParameter
    
        ' Unit ID
        Set adoParameter = adoCommand.CreateParameter("UNITID", adBSTR, adParamInput, Len(strUnitId))
        adoParameter.Attributes = adParamNullable
        If Len(strUnitId) > 0 Then
            adoParameter.Value = strUnitId
        Else
            adoParameter.Value = Null
        End If
        adoCommand.Parameters.Append adoParameter
    
        ' User ID
        Set adoParameter = adoCommand.CreateParameter("USERID", adBSTR, adParamInput, Len(strUserId))
        adoParameter.Attributes = adParamNullable
        If Len(strUserId) > 0 Then
            adoParameter.Value = strUserId
        Else
            adoParameter.Value = Null
        End If
        adoCommand.Parameters.Append adoParameter
    
        ' Mortgage Product Code
        Set adoParameter = adoCommand.CreateParameter("MORTGAGEPRODUCTCODE", adBSTR, adParamInput, Len(strMortgageProductCode))
        adoParameter.Attributes = adParamNullable
        If Len(strMortgageProductCode) > 0 Then
            adoParameter.Value = strMortgageProductCode
        Else
            adoParameter.Value = Null
        End If
        adoCommand.Parameters.Append adoParameter
    
        ' Mortgage Product Start Date
        Set adoParameter = adoCommand.CreateParameter("MORTGAGEPRODUCTSTARTDATE", adDate, adParamInput)
        adoParameter.Attributes = adParamNullable
        If Len(strMortgageProductStartDate) > 0 Then
            adoParameter.Value = objConvertAssist.CSafeDate(strMortgageProductStartDate)
        Else
            adoParameter.Value = Null
        End If
        adoCommand.Parameters.Append adoParameter
    
        ' Type Of Application
        Set adoParameter = adoCommand.CreateParameter("TYPEOFAPPLICATION", adInteger, adParamInput)
        adoParameter.Attributes = adParamNullable
        If Len(strTypeOfApplication) > 0 Then
            adoParameter.Value = objConvertAssist.CSafeInt(strTypeOfApplication)
        Else
            adoParameter.Value = Null
        End If
        adoCommand.Parameters.Append adoParameter
    
        'Application From Stage
        Set adoParameter = adoCommand.CreateParameter("APPLICATIONSTAGEFROM", adInteger, adParamInput)
        adoParameter.Attributes = adParamNullable
        If Len(strApplicationFromStage) > 0 Then
            adoParameter.Value = objConvertAssist.CSafeInt(strApplicationFromStage)
        Else
            adoParameter.Value = Null
        End If
        adoCommand.Parameters.Append adoParameter
    
        'Application To Stage
        Set adoParameter = adoCommand.CreateParameter("APPLICATIONSTAGETO", adInteger, adParamInput)
        adoParameter.Attributes = adParamNullable
        If Len(strApplicationToStage) > 0 Then
            adoParameter.Value = objConvertAssist.CSafeInt(strApplicationToStage)
        Else
            adoParameter.Value = Null
        End If
        adoCommand.Parameters.Append adoParameter
    
        ' Application Date From
        Set adoParameter = adoCommand.CreateParameter("APPLICATIONDATEFROM", adDate, adParamInput)
        adoParameter.Attributes = adParamNullable
        If Len(strApplicationDateFrom) > 0 Then
            adoParameter.Value = objConvertAssist.CSafeDate(strApplicationDateFrom)
        Else
            adoParameter.Value = Null
        End If
        adoCommand.Parameters.Append adoParameter
    
        ' Application Date To
        Set adoParameter = adoCommand.CreateParameter("APPLICATIONDATETO", adDate, adParamInput)
        adoParameter.Attributes = adParamNullable
        If Len(strApplicationDateTo) > 0 Then
            adoParameter.Value = objConvertAssist.CSafeDate(strApplicationDateTo)
        Else
            adoParameter.Value = Null
        End If
        adoCommand.Parameters.Append adoParameter
    
        ' Approval Date
        Set adoParameter = adoCommand.CreateParameter("APPLICATIONAPPROVALDATE", adDate, adParamInput)
        adoParameter.Attributes = adParamNullable
        If Len(strApprovalYear) > 0 Then
            'BM0523 Changed APPLICATIONAPPROVALMONTH to APPLICATIONAPPROVEDMONTH
            strApprovalMonth = m_objXmlAssist.GetTagValue(vobjXmlRequest, "APPLICATIONAPPROVEDMONTH")
            'BM0523 End
            If Len(strApprovalMonth) > 0 Then
                adoParameter.Value = objConvertAssist.CSafeDate(DateSerial(strApprovalYear, strApprovalMonth, 1))
            Else
                adoParameter.Value = Null
            End If
        Else
            adoParameter.Value = Null
        End If
        adoCommand.Parameters.Append adoParameter
        
        'Introducer ID
        Set adoParameter = adoCommand.CreateParameter("INTRODUCERID", adBSTR, adParamInput, Len(strIntroducerID))
        adoParameter.Attributes = adParamNullable
        If Len(strIntroducerID) > 0 Then
            adoParameter.Value = strIntroducerID
        Else
            adoParameter.Value = Null
        End If
        adoCommand.Parameters.Append adoParameter
        
        'Include Cancelled Apps
        Set adoParameter = adoCommand.CreateParameter("INCLUDECANCELLEDAPPS", adBSTR, adParamInput, Len(strIncludeCancelledApps))
        adoParameter.Attributes = adParamNullable
        If Len(strIncludeCancelledApps) > 0 Then
            adoParameter.Value = strIncludeCancelledApps
        Else
            adoParameter.Value = Null
        End If
        adoCommand.Parameters.Append adoParameter
    
        ' Include Declined Apps
        Set adoParameter = adoCommand.CreateParameter("INCLUDEDECLINEDAPPS", adBSTR, adParamInput, Len(strIncludeDeclinedApps))
        adoParameter.Attributes = adParamNullable
        If Len(strIncludeDeclinedApps) > 0 Then
            adoParameter.Value = strIncludeDeclinedApps
        Else
            adoParameter.Value = Null
        End If
        adoCommand.Parameters.Append adoParameter
        
        'MO - 01/10/2002 - BMIDS00502, INWP1 - Start
        
        ' Direct Indirect Business
        Set adoParameter = adoCommand.CreateParameter("DIRECTINDIRECTBUSINESS", adInteger, adParamInput)
        adoParameter.Attributes = adParamNullable
        If Len(strDirectIndirectBusiness) > 0 Then
            adoParameter.Value = objConvertAssist.CSafeInt(strDirectIndirectBusiness)
        Else
            adoParameter.Value = Null
        End If
        adoCommand.Parameters.Append adoParameter
        
        'MO - 01/10/2002 - BMIDS00502, INWP1 - End
        
        'BM0523 Searching by customer details is now done in FindApplicationListByCustomer
        
        ' Postcode
        Set adoParameter = adoCommand.CreateParameter("POSTCODE", adBSTR, adParamInput, Len(strPostcode))
        adoParameter.Attributes = adParamNullable
        If Len(strPostcode) > 0 Then
            adoParameter.Value = strPostcode
        Else
            adoParameter.Value = Null
        End If
        adoCommand.Parameters.Append adoParameter
        
        'MAR46
        'Include Product Switch Completed Applications
        Set adoParameter = adoCommand.CreateParameter("INCLUDEPRODUCTSWITCHAPPS", adBSTR, adParamInput, Len(strIncludeProductSwitchApps))
        adoParameter.Attributes = adParamNullable
        If Len(strIncludeProductSwitchApps) > 0 Then
            adoParameter.Value = strIncludeProductSwitchApps
        Else
            adoParameter.Value = Null
        End If
        adoCommand.Parameters.Append adoParameter

        'Include Transfer of Equity Completed Applications
        Set adoParameter = adoCommand.CreateParameter("INCLUDETOEAPPS", adBSTR, adParamInput, Len(strIncludeToEApps))
        adoParameter.Attributes = adParamNullable
        If Len(strIncludeToEApps) > 0 Then
            adoParameter.Value = strIncludeToEApps
        Else
            adoParameter.Value = Null
        End If
        adoCommand.Parameters.Append adoParameter
        
        'BM0525
        ' Total number of records (output parameter)
        Set adoTotalRecordsParam = adoCommand.CreateParameter("@p_TOTALRECORDS", adInteger, adParamOutput)
        adoCommand.Parameters.Append adoTotalRecordsParam
        'BM0525 End
    
        'BM0523 Changed to use SQL XML to get XML direct from the database for better performance
        Set adoStream = New ADODB.Stream
        adoStream.open
        
        With adoCommand
            Set .ActiveConnection = adoConnection
            .Properties("Output Stream") = adoStream
            .Properties("XML Root") = "APPLICATIONLIST"
            .Execute , , adExecuteStream
            Set .ActiveConnection = Nothing
        End With
            
        Set xmlDoc = New FreeThreadedDOMDocument40
        xmlDoc.validateOnParse = False
        xmlDoc.setProperty "NewParser", True
        
        If xmlDoc.loadXML(adoStream.ReadText) Then
        'Read the output parameters to get the totals
            lngTotalRecords = adoTotalRecordsParam.Value
            If lngTotalRecords = 0 Then
                m_objErrAssist.ThrowError TypeName(Me), cstrFunctionName, oeRecordNotFound
            End If
        
            xmlDoc.documentElement.setAttribute "TOTALRECORDS", lngTotalRecords
            
            ' Convert the XML returned from the database to the expected format
            For Each xmlNode In xmlDoc.selectNodes("APPLICATIONLIST/APPLICATION")
                Set xmlTypeOfApp = xmlNode.selectSingleNode("TYPEOFAPPLICATION")
                Set xmlTypeOfAppText = xmlNode.selectSingleNode("TYPEOFAPPLICATIONTEXT")
                xmlTypeOfApp.setAttribute "TEXT", xmlTypeOfAppText.Text
                xmlNode.removeChild xmlTypeOfAppText
            Next
        Else
            ' raise application error to be interpreted by calling object
            m_objErrAssist.ThrowError TypeName(Me), cstrFunctionName, oeRecordNotFound
        End If
        'BM0523 End
        
        adoConnection.Close
        Set adoConnection = Nothing
        
        Set IApplicationManagerDO_FindApplicationList = xmlDoc.documentElement

        Set adoCommand = Nothing
        Set objADOAssist = Nothing
        Set objSqlAssist = Nothing
        Set objConvertAssist = Nothing
        Set adoParameter = Nothing
        Set xmlNode = Nothing
        Set xmlDoc = Nothing
        Set xmlTypeOfApp = Nothing
        Set xmlTypeOfAppText = Nothing
        Set adoTotalRecordsParam = Nothing

    End If
    
    Exit Function

FindApplicationListVbErr:
    
    ' Close the connection if still open
    If Not adoConnection Is Nothing Then
        If adoConnection.State = adStateOpen Then
            adoConnection.Close
        End If
    End If

    Set objADOAssist = Nothing
    Set objSqlAssist = Nothing
    Set objConvertAssist = Nothing
    Set adoConnection = Nothing
    Set adoCommand = Nothing
    Set adoParameter = Nothing
    Set xmlNode = Nothing
    Set xmlDoc = Nothing
    Set xmlTypeOfApp = Nothing
    Set xmlTypeOfAppText = Nothing
    Set adoTotalRecordsParam = Nothing
    
    '   re-raise error for business object to interpret as appropriate
    Debug.Print Err.Number & ": " & Err.Description
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

'BM0523 GHun 22/05/2003
Private Function FindApplicationListByCustomer(ByVal vobjXmlRequest As IXMLDOMNode) As IXMLDOMElement
                                        
    Const cstrFunctionName As String = "FindApplicationListByCustomer"
    
    Dim objADOAssist            As ADOAssist
    Dim objSqlAssist            As SQLAssist
    Dim adoConnection           As ADODB.Connection
    Dim adoCommand              As ADODB.Command
    Dim adoParameter            As ADODB.Parameter
    Dim adoTotalRecordsParam    As ADODB.Parameter
    Dim adoStream               As ADODB.Stream
    Dim xmlDoc                  As FreeThreadedDOMDocument40
    Dim xmlNode                 As IXMLDOMNode
    Dim xmlTypeOfAppText        As IXMLDOMNode
    Dim xmlTypeOfApp            As IXMLDOMElement
    Dim lngStartRow             As Long
    Dim lngNumRows              As Long
    Dim lngTotalRecords         As Long
    Dim strSQL                  As String
    Dim strIncludeCancelledApps As String
    Dim strIncludeDeclinedApps  As String
    Dim strSurname              As String
    Dim strFirstForename        As String
    Dim strSecondForename       As String
    Dim strOtherForenames       As String
    Dim strPostcode             As String
    Dim strDateOfBirth          As String
    Dim blnIsWildCard           As Boolean
    
    'MAR46
    Dim strIncludeProductSwitchApps As String
    Dim strIncludeToEApps           As String
  
        
On Error GoTo FindApplicationListByCustomerVbErr
    
    Set objADOAssist = New ADOAssist
    Set objSqlAssist = New SQLAssist
    Set adoConnection = New ADODB.Connection
    Set adoCommand = New ADODB.Command
               
    Set xmlNode = vobjXmlRequest.selectSingleNode("//REQUEST")
       
    'Determine the start record and number of records to return
    GetRecordRange xmlNode.cloneNode(False), lngStartRow, lngNumRows
      
    strIncludeCancelledApps = m_objXmlAssist.GetTagValue(vobjXmlRequest, "INCLUDECANCELLEDAPPS")
    strIncludeDeclinedApps = m_objXmlAssist.GetTagValue(vobjXmlRequest, "INCLUDEDECLINEDAPPS")
    strSurname = m_objXmlAssist.GetTagValue(vobjXmlRequest, "SURNAME")
    strSurname = objSqlAssist.FormatWildcardedString(strSurname, blnIsWildCard, True)
    strFirstForename = m_objXmlAssist.GetTagValue(vobjXmlRequest, "FIRSTFORENAME")
    strFirstForename = objSqlAssist.FormatWildcardedString(strFirstForename, blnIsWildCard, True)
    strSecondForename = m_objXmlAssist.GetTagValue(vobjXmlRequest, "SECONDFORENAME")
    strSecondForename = objSqlAssist.FormatWildcardedString(strSecondForename, blnIsWildCard, True)
    If Len(strSecondForename) = 0 Then
        strSecondForename = "%"
    End If
    strOtherForenames = m_objXmlAssist.GetTagValue(vobjXmlRequest, "OTHERFORENAMES")
    strOtherForenames = objSqlAssist.FormatWildcardedString(strOtherForenames, blnIsWildCard, True)
    If Len(strOtherForenames) = 0 Then
        strOtherForenames = "%"
    End If
    strPostcode = m_objXmlAssist.GetTagValue(vobjXmlRequest, "POSTCODE")
    strPostcode = objSqlAssist.FormatWildcardedString(strPostcode, blnIsWildCard, True)
    If Len(strPostcode) = 0 Then
        strPostcode = "%"
    End If
    strDateOfBirth = m_objXmlAssist.GetTagValue(vobjXmlRequest, "DATEOFBIRTH")
    If Len(strDateOfBirth) = 0 Then
        strDateOfBirth = "%"
    End If

    'MAR46
    strIncludeProductSwitchApps = m_objXmlAssist.GetTagValue(vobjXmlRequest, "INCPRODUCTSWITCHAPPSCHECKED")
    strIncludeToEApps = m_objXmlAssist.GetTagValue(vobjXmlRequest, "INCTOEAPPSCHECKED")

    With adoConnection
        .ConnectionString = objADOAssist.GetConnStr
        .CursorLocation = adUseClient
        .open
    End With

#If GENERIC_SQL Then
    Select Case objADOAssist.GetDBEngine
    Case SQLServer
        strSQL = "USP_FINDAPPLICATIONLISTBYCUSTOMER"
        adoCommand.CommandType = adCmdStoredProc
    Case Oracle
        ' Not implemented
    End Select
#Else
    ' Not implemented
    adoCommand.CommandType = adCmdText
#End If

    adoCommand.CommandText = strSQL

    ' Start Record
    Set adoParameter = adoCommand.CreateParameter(, adInteger, adParamInput, , lngStartRow)
    adoCommand.Parameters.Append adoParameter

    ' Maximum number of records to return
    Set adoParameter = adoCommand.CreateParameter(, adInteger, adParamInput, , lngNumRows)
    adoCommand.Parameters.Append adoParameter
       
    'Include Cancelled Apps
    Set adoParameter = adoCommand.CreateParameter("INCLUDECANCELLEDAPPS", adBSTR, adParamInput, Len(strIncludeCancelledApps))
    adoParameter.Attributes = adParamNullable
    If Len(strIncludeCancelledApps) > 0 Then
        adoParameter.Value = strIncludeCancelledApps
    Else
        adoParameter.Value = Null
    End If
    adoCommand.Parameters.Append adoParameter

    ' Include Declined Apps
    Set adoParameter = adoCommand.CreateParameter("INCLUDEDECLINEDAPPS", adBSTR, adParamInput, Len(strIncludeDeclinedApps))
    adoParameter.Attributes = adParamNullable
    If Len(strIncludeDeclinedApps) > 0 Then
        adoParameter.Value = strIncludeDeclinedApps
    Else
        adoParameter.Value = Null
    End If
    adoCommand.Parameters.Append adoParameter
       
    ' Surname
    Set adoParameter = adoCommand.CreateParameter("SURNAME", adBSTR, adParamInput, Len(strSurname))
    adoParameter.Attributes = adParamNullable
    If Len(strSurname) > 0 Then
        adoParameter.Value = strSurname
    Else
        adoParameter.Value = Null
    End If
    adoCommand.Parameters.Append adoParameter
    
    ' First Forename
    Set adoParameter = adoCommand.CreateParameter("FIRSTFORENAME", adBSTR, adParamInput, Len(strFirstForename))
    adoParameter.Attributes = adParamNullable
    If Len(strFirstForename) > 0 Then
        adoParameter.Value = strFirstForename
    Else
        adoParameter.Value = Null
    End If
    adoCommand.Parameters.Append adoParameter
    
    ' Second Forename
    Set adoParameter = adoCommand.CreateParameter("SECONDFORENAME", adBSTR, adParamInput, Len(strSecondForename))
    adoParameter.Attributes = adParamNullable
    If Len(strSecondForename) > 0 Then
        adoParameter.Value = strSecondForename
    Else
        adoParameter.Value = Null
    End If
    adoCommand.Parameters.Append adoParameter
    
    ' Other Forenames
    Set adoParameter = adoCommand.CreateParameter("OTHERFORENAMES", adBSTR, adParamInput, Len(strOtherForenames))
    adoParameter.Attributes = adParamNullable
    If Len(strOtherForenames) > 0 Then
        adoParameter.Value = strOtherForenames
    Else
        adoParameter.Value = Null
    End If
    adoCommand.Parameters.Append adoParameter
    
    ' Postcode
    Set adoParameter = adoCommand.CreateParameter("POSTCODE", adBSTR, adParamInput, Len(strPostcode))
    adoParameter.Attributes = adParamNullable
    If Len(strPostcode) > 0 Then
        adoParameter.Value = strPostcode
    Else
        adoParameter.Value = Null
    End If
    adoCommand.Parameters.Append adoParameter
    
    ' Date of Birth
    Set adoParameter = adoCommand.CreateParameter("DATEOFBIRTH", adBSTR, adParamInput, Len(strDateOfBirth))
    adoParameter.Attributes = adParamNullable
    If Len(strDateOfBirth) > 0 Then
        adoParameter.Value = strDateOfBirth
    Else
        adoParameter.Value = Null
    End If
    adoCommand.Parameters.Append adoParameter

    'MAR46  Include Product Switch Completed Applications
    Set adoParameter = adoCommand.CreateParameter("INCLUDEPRODUCTSWITCHAPPS", adBSTR, adParamInput, Len(strIncludeProductSwitchApps))
    adoParameter.Attributes = adParamNullable
    If Len(strIncludeProductSwitchApps) > 0 Then
        adoParameter.Value = strIncludeProductSwitchApps
    Else
        adoParameter.Value = Null
    End If
    adoCommand.Parameters.Append adoParameter

    'MAR46  Include Transfer of Equity Completed Applications
    Set adoParameter = adoCommand.CreateParameter("INCLUDETOEAPPS", adBSTR, adParamInput, Len(strIncludeToEApps))
    adoParameter.Attributes = adParamNullable
    If Len(strIncludeToEApps) > 0 Then
        adoParameter.Value = strIncludeToEApps
    Else
        adoParameter.Value = Null
    End If
    adoCommand.Parameters.Append adoParameter

    ' Total number of records found (output parameter)
    Set adoTotalRecordsParam = adoCommand.CreateParameter("@p_TOTALRECORDS", adInteger, adParamOutput)
    adoCommand.Parameters.Append adoTotalRecordsParam
    
    Set adoStream = New ADODB.Stream
    adoStream.open
    
    With adoCommand
        Set .ActiveConnection = adoConnection
        .Properties("Output Stream") = adoStream
        .Properties("XML Root") = "APPLICATIONLIST"
        .Execute , , adExecuteStream
        Set .ActiveConnection = Nothing
    End With
        
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    
    If xmlDoc.loadXML(adoStream.ReadText) Then
        'Read the output parameters to get the totals
        lngTotalRecords = adoTotalRecordsParam.Value
        If lngTotalRecords = 0 Then
            m_objErrAssist.ThrowError TypeName(Me), cstrFunctionName, oeRecordNotFound
        End If
            
        xmlDoc.documentElement.setAttribute "TOTALRECORDS", lngTotalRecords
        
        ' Convert the XML returned from the database to the expected format
        For Each xmlNode In xmlDoc.selectNodes("APPLICATIONLIST/APPLICATION")
            Set xmlTypeOfApp = xmlNode.selectSingleNode("TYPEOFAPPLICATION")
            Set xmlTypeOfAppText = xmlNode.selectSingleNode("TYPEOFAPPLICATIONTEXT")
            xmlTypeOfApp.setAttribute "TEXT", xmlTypeOfAppText.Text
            xmlNode.removeChild xmlTypeOfAppText
        Next
    Else
        ' raise application error to be interpreted by calling object
        m_objErrAssist.ThrowError TypeName(Me), cstrFunctionName, oeRecordNotFound
    End If
    
    adoConnection.Close
    Set adoConnection = Nothing
    
    Set FindApplicationListByCustomer = xmlDoc.documentElement
    
    Set adoCommand = Nothing
    Set objADOAssist = Nothing
    Set objSqlAssist = Nothing
    Set adoParameter = Nothing
    Set xmlNode = Nothing
    Set xmlDoc = Nothing
    Set xmlTypeOfApp = Nothing
    Set xmlTypeOfAppText = Nothing
    Set adoTotalRecordsParam = Nothing
    Set adoStream = Nothing

    Exit Function

FindApplicationListByCustomerVbErr:
    
    ' Close the connection if still open
    If Not adoConnection Is Nothing Then
        If adoConnection.State = adStateOpen Then
            adoConnection.Close
        End If
    End If

    Set objADOAssist = Nothing
    Set objSqlAssist = Nothing
    Set adoConnection = Nothing
    Set adoCommand = Nothing
    Set adoParameter = Nothing
    Set xmlNode = Nothing
    Set xmlDoc = Nothing
    Set xmlTypeOfApp = Nothing
    Set xmlTypeOfAppText = Nothing
    Set adoTotalRecordsParam = Nothing
    Set adoStream = Nothing
    
    '   re-raise error for business object to interpret as appropriate
    Debug.Print Err.Number & ": " & Err.Description
    Err.Raise Err.Number, Err.Source, Err.Description

End Function
'BM0523 End

'BM0523 GHun 22/05/2003
Private Function FindApplicationListByAppNum(ByVal vobjXmlRequest As IXMLDOMNode) As IXMLDOMElement
                                        
    Const cstrFunctionName As String = "FindApplicationListByAppNum"
    
    Dim objADOAssist            As ADOAssist
    Dim adoConnection           As ADODB.Connection
    Dim adoCommand              As ADODB.Command
    Dim adoParameter            As ADODB.Parameter
    Dim adoStream               As ADODB.Stream
    Dim xmlDoc                  As FreeThreadedDOMDocument40
    Dim strSQL                  As String
    Dim strApplicationNumber    As String
    Dim xmlTypeOfAppText        As IXMLDOMNode       'BMIDS607
    Dim xmlTypeOfApp            As IXMLDOMElement    'BMIDS607
    Dim xmlNode                 As IXMLDOMNode       'BMIDS607
   
        
On Error GoTo FindApplicationListByAppNumVbErr
    
    Set objADOAssist = New ADOAssist
    Set adoConnection = New ADODB.Connection
    Set adoCommand = New ADODB.Command
            
    strApplicationNumber = m_objXmlAssist.GetTagValue(vobjXmlRequest, "APPLICATIONNUMBER")

    With adoConnection
        .ConnectionString = objADOAssist.GetConnStr
        .CursorLocation = adUseClient
        .open
    End With

#If GENERIC_SQL Then
    Select Case objADOAssist.GetDBEngine
    Case SQLServer
        strSQL = "USP_FINDAPPLICATIONLISTBYAPPNUM"
        adoCommand.CommandType = adCmdStoredProc
    Case Oracle
        ' Not implemented
    End Select
#Else
    ' Not implemented
    adoCommand.CommandType = adCmdText
#End If

    adoCommand.CommandText = strSQL
       
    ' Application Number
    Set adoParameter = adoCommand.CreateParameter("APPLICATIONNUMBER", adBSTR, adParamInput, Len(strApplicationNumber))
    adoParameter.Attributes = adParamNullable
    If Len(strApplicationNumber) > 0 Then
        adoParameter.Value = strApplicationNumber
    Else
        adoParameter.Value = Null
    End If
    adoCommand.Parameters.Append adoParameter
    
    Set adoStream = New ADODB.Stream
    adoStream.open
    
    With adoCommand
        .ActiveConnection = adoConnection
        .Properties("Output Stream") = adoStream
        .Properties("XML Root") = "APPLICATIONLIST"
        .Execute , , adExecuteStream
        Set .ActiveConnection = Nothing
    End With
        
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    
    If xmlDoc.loadXML(adoStream.ReadText) Then
        
        If xmlDoc.selectSingleNode("APPLICATIONLIST/APPLICATION") Is Nothing Then
            m_objErrAssist.ThrowError TypeName(Me), cstrFunctionName, oeRecordNotFound
        End If
           
        xmlDoc.documentElement.setAttribute "TOTALRECORDS", 1
        
        'BMIDS607  Convert the XML returned from the database to the expected format
        Set xmlNode = xmlDoc.selectSingleNode("APPLICATIONLIST/APPLICATION")
        Set xmlTypeOfApp = xmlNode.selectSingleNode("TYPEOFAPPLICATION")
        Set xmlTypeOfAppText = xmlNode.selectSingleNode("TYPEOFAPPLICATIONTEXT")
        xmlTypeOfApp.setAttribute "TEXT", xmlTypeOfAppText.Text
        xmlNode.removeChild xmlTypeOfAppText
        
    Else
        ' raise application error to be interpreted by calling object
        m_objErrAssist.ThrowError TypeName(Me), cstrFunctionName, oeRecordNotFound
    End If
    
    adoConnection.Close
    Set adoConnection = Nothing
    
    Set FindApplicationListByAppNum = xmlDoc.documentElement
    
    Set objADOAssist = Nothing
    Set adoCommand = Nothing
    Set adoParameter = Nothing
    Set xmlDoc = Nothing
    Set adoStream = Nothing

    Exit Function

FindApplicationListByAppNumVbErr:
    
    ' Close the connection if still open
    If Not adoConnection Is Nothing Then
        If adoConnection.State = adStateOpen Then
            adoConnection.Close
        End If
    End If

    Set objADOAssist = Nothing
    Set adoConnection = Nothing
    Set adoCommand = Nothing
    Set adoParameter = Nothing
    Set xmlDoc = Nothing
    Set adoStream = Nothing
    
    '   re-raise error for business object to interpret as appropriate
    Debug.Print Err.Number & ": " & Err.Description
    Err.Raise Err.Number, Err.Source, Err.Description

End Function
'BM0523 End

'BM0523 GHun 23/05/2003
Private Sub GetRecordRange(ByVal vxmlNode As IXMLDOMNode, ByRef lngStartRow As Long, lngPageSize As Long)
    Dim strPageSize As String
    Dim strPageNum  As String
    Dim lngPageNum  As Long
    
    lngStartRow = 1
    lngPageSize = cintFINDAPPLICATIONLISTPAGE
    
    If Not vxmlNode Is Nothing Then
        
        If Not vxmlNode.Attributes.getNamedItem("PAGESIZE") Is Nothing Then
            strPageSize = vxmlNode.Attributes.getNamedItem("PAGESIZE").Text
            
            If Len(strPageSize) > 0 Then
                lngPageSize = CLng(strPageSize)
            End If
        End If
        
        If lngPageSize < 1 Then
            lngPageSize = cintFINDAPPLICATIONLISTPAGE
        End If
        
        If Not vxmlNode.Attributes.getNamedItem("PAGENUMBER") Is Nothing Then
            strPageNum = vxmlNode.Attributes.getNamedItem("PAGENUMBER").Text
            
            If Len(strPageNum) > 0 Then
                lngPageNum = CLng(strPageNum)
                If lngPageNum < 1 Then
                    lngPageNum = 1
                End If
                lngStartRow = ((lngPageNum - 1) * lngPageSize) + 1
            End If
        End If
    
    End If
End Sub
'BM0523 End

'BMIDS758  Add function to create an entry in the RemovedToECustomer table

Private Sub IApplicationManagerDO_CreateRemovedToECustomer(ByVal vxmlTableElement As IXMLDOMElement)
' header -----------------------------------------------------------------------
' description:
'   create an instance of the persistant data associated with this data object
'   for each set of data in the request
' pass:
'   vxmlTableElement  xml element containing the request
' return:                n/a
'--------------------------------------------------------------------------------
On Error GoTo CreateVbErr

    Const strFunctionName As String = "IApplicationManagerDO_CreateRemovedToECustomer"
   
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
    
    'EP2_56 GHun
    Dim xmlNode As IXMLDOMElement
    Dim lngSequenceNumber As Long
    Dim xmlTOECustomer  As IXMLDOMElement
   
    Set xmlTOECustomer = vxmlTableElement.selectSingleNode(".//REMOVEDTOECUSTOMER")
    If xmlTOECustomer Is Nothing Then
        Set xmlTOECustomer = vxmlTableElement
    End If
    'EP2_56 End
    
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadRemovedToECustomer
    
    'EP2_56 GHun Get the next highest sequence number
    lngSequenceNumber = m_objIDOAssist.GetNextSequenceNumber(vxmlTableElement, xmlClassDefDoc, "REMOVEDTOECUSTOMER", "SEQUENCENUMBER")
      
    Set xmlNode = vxmlTableElement.selectSingleNode(".//SEQUENCENUMBER")
    
    If Not xmlNode Is Nothing Then
        xmlNode.Text = CStr(lngSequenceNumber)
    Else
        Set xmlNode = vxmlTableElement.ownerDocument.createElement("SEQUENCENUMBER")
        xmlNode.Text = CStr(lngSequenceNumber)
        xmlTOECustomer.appendChild xmlNode
    End If
    'EP2_56 End
    
    m_objIDOAssist.Create vxmlTableElement, xmlClassDefDoc
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set xmlNode = Nothing 'EP2_56 GHun
    Set xmlTOECustomer = Nothing    'EP2_56 GHun

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

    Exit Sub

CreateVbErr:
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set xmlNode = Nothing 'EP2_56 GHun
    Set xmlTOECustomer = Nothing    'EP2_56 GHun
    
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

