VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "ApplicationDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'Workfile:      ApplicationDO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling

'Description:   Code template for Application Data Object
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
'Prog   Date        Description
'AD     05/11/99    Created.
'RF     10/01/99    Added methods for the Declaration table.
'SR     13/03/00    New function AddCurrentDateAndTimeToRequest. This is used in the methods
'                   'CreateApplicationStage' and 'UpdateApplicationStage'
' IW    16/03/00    Added LoadApplicationVerificationData
' IW    19/03/00    Added LoadApplicationCreditCardData
'SR     24/03/00    Added MemoPadMethods
'MC     26/04/00    Added OtherInsuranceCompany methods
'MC     03/05/00    Fixed SYS0662
'SR     02/06/00    Added methods 'CreateNewLoanDetails' and 'UpdateNewLoanDetails'
'BG     02/06/00    Added Method 'GetNewLoan'
'MC     13/07/00    SYS1168 - Update DeleteDependantsForCustomers
'LD     07/11/00    Explicity close database connections
'LD     07/11/00    Explicity close recordsets
'DJP    19/12/00    Change CreateUserHistory to auto generate the UserHistoryDate
'DJP    19/12/00    Make CreateUserHistory create it's own domdocument as it can fail
'                   under certain circumstances
'JLD    02/01/01    SYS1739 removed change to FindListMultiple which included an order direction
'                   and put the order direction in with the orderBy string.
'ASm    12/01/01    SYS1817: oeWasDependant, oeWasOtherResident error enums removed and
'                   replaced by their respective error numbers
'DJP    29/01/01    SYS1808 changed GetCurrentApplicationStage to match on CASESEQUENCENO
'APS    27/02/01    SYS1990 changed GetCurrentApplicationStage to match on CASESEQUENCENO
'APS    27/02/01    SYS1993 Added CreateApplicationPriority  and amended GetApplicationPriority methods
'PSC    08/03/01    SYS1879 Added GetMortgageAccountData
'JLD    08/03/01    SYS1879 check for null node in CreateApplicationPriority
'APS    11/03/01    SYS2030 Added check to CreateApplicationStage processing
'ASt    12/03/01    CC012 Now compatible with SQL Server and Oracle
'JR     22/03/01    SYS2048 Added GetApplicationOfferData
'JR     20/04/01    SYS2048 Commented out objContext.setabort/complete in GetApplicationData/GetApplicationOfferData methods
'LD     11/06/01    SYS2367 SQL Server Port - Length must be specified in calls to CreateParameter
'SR     20/07/01  SYS2412 New method 'FindApplicationFactFindList'
'BG     21/12/01    SYS3268 Added UpdateApplicationOffer and CreateApplicationOffer methods
'------------------------------------------------------------------------------------------
'------------------------------------------------------------------------------------------
'BMIDS Specific History:
'
'Prog   Date        AQR         Description
'MV     12/08/2002  BMIDS00319  Core AQR Ref: SYS1730 ; Modified IApplicationDO_DeleteOtherResident
'                               IApplicationDO_DeleteDependantForCustomer
'MV     13/08/2002  BMIDS0331   Core Ref AQR : SYS 2372 Amended GetNextApplicationNumber
'SR     02/01/2003  BM0209      Modified GetApplicationData (cope with new tag to request 'DataForCMOnly')
'PSC    09/01/2003  BM0089      Amend GetCurrentApplicationStage to use No Lock hint
'GD     27/01/2003  BM0275      Amend Findmemopadlist to use validation types, and not value ids.
'LD     02/01/2003  BM0391      Use COUNT(1) rather than COUNT(*)
'MC     11/06/2004  BMIDS763    ApplicationDate column included in SQL Query.
'SR     13/08/2004  BMIDS815    Modified method 'GetApplicationData'
'KRW    26/08/2004  BMIDS810    Added extra conditioning in GetApplicationOfferData to pick up highest OFFERSEQUENCENO
'KRW    22/11/2004  BMIDS888    GetApplicationData change for Long Running Threads and Com+ errors on O4 App Servers
'KRW    02/12/2004  BMIDS888    GetApplicationData ammendments
'JD     15/03/2005  BMIDS975    Added FurtherAdvanceTermInd to calcs return of GetApplicationData
'------------------------------------------------------------------------------------------
'MARS Specific History:
'HMA    25/07/2005  MAR18       Add LevelOfAdvice to GetApplicationData
'MF     19/08/2005  MAR20       Added CretePropertyDetails
'JD     26/09/2005  MAR40       Added save of leasehold details to CreatePropertyDetails and updateloanproperty
'SR     20/10/2005  MAR24       New methods 'GetApplicationUnderwriting', 'CreateApplicationUnderwriting'
'                               and 'UpdateApplicationUnderwriting'
'SR     20/10/2005  MAR253      Modifed IApplicationDO_GetApplicationUnderwriting
'MV     22/10/2005  MAR266      Amended IApplicationDO_CreatePropertyDetails
'GHun   24/10/2005  MAR109      Changed CreatePropertyDetails
'------------------------------------------------------------------------------------------
'EPSOM Specific History:
'IK    10/04/2006   EP241       fix MAR40 change to UpdateLoanProperty
'                               not fixed by MAR109
'                               NEWPROPERTYLEASEHOLD is not mandatory
'IK     17/05/2006  EP561       add IsCheckDigitRequired test
'PB     15/08/2006  EP1089      Merge MAR1890 - added GetDecisionDetails
'INR    25/01/2007  EP2_780     return of SPECIALSCHEME from getapplicationdata
'------------------------------------------------------------------------------------------
Option Explicit
Private Const cstrTABLE_NAME = "APPLICATION"
Implements IApplicationDO
Implements ObjectControl

Private m_objContext As ObjectContext
Private m_objErrAssist As ErrAssist
Private m_objXmlAssist As XMLAssist
Private m_objIDOAssist As IDOAssist

' header ----------------------------------------------------------------------------------
' description:
'   Create an instance of the persistant data associated with this data object
'   for each set of data in the request
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
Private Sub IApplicationDO_Create(ByVal vxmlTableElement As IXMLDOMElement)
    
    On Error GoTo CreateVbErr
    Const strFunctionName As String = "IApplicationDO_Create"
   
    Dim objErrAssist As New ErrAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
    Dim objIDOAssist As IDOAssist
    Dim objContext As ObjectContext
    
    Set objContext = GetObjectContext()
                    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadApplicationData
   
   ' application date may requre setting
    AddDateToRequest vxmlTableElement

    objIDOAssist.Create vxmlTableElement, xmlClassDefDoc
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Sub

CreateVbErr:
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
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
End Sub

Private Function IApplicationDO_GetApplicationUnderwriting(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode

On Error GoTo GetApplicationUnderwritingVbErr

    Const strFunctionName As String = "IApplicationDO_GetApplicationUnderwriting"
    
    Dim objErrAssist As New ErrAssist
    
    Dim objIDOAssist As IDOAssist, _
        objXmlAssist As New XMLAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
    Dim xmlReturn As IXMLDOMNode
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
            
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadApplicationUnderwritingData
    Set xmlReturn = objIDOAssist.GetData(vxmlTableElement, xmlClassDefDoc)

    Set IApplicationDO_GetApplicationUnderwriting = xmlReturn

    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlReturn = Nothing
    Set objXmlAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

GetApplicationUnderwritingVbErr:
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set xmlReturn = Nothing
    Set objXmlAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
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

Private Function IApplicationDO_GetDecisionDetails(ByVal vxmlTableElement As IXMLDOMElement) As String
' header ----------------------------------------------------------------------------------
' description:
'   calls stored procedure to return decision details
'------------------------------------------------------------------------------------------
On Error GoTo GetDecisionDetailsErr

    Const cstrFunctionName As String = "IApplicationDO_GetDecisionDetails"
       
    Dim sApplicationNumber As String
    Dim sApplicationFactFindNumber As String
    Dim oCommand As ADODB.Command
    Dim adoConn As ADODB.Connection
    Dim adoStream As ADODB.Stream
   
    'Use XML assist and obtain the parameters to pass into the command object.
    sApplicationNumber = xmlGetMandatoryNodeText(vxmlTableElement, "//APPLICATIONNUMBER")
    sApplicationFactFindNumber = xmlGetMandatoryNodeText(vxmlTableElement, "//APPLICATIONFACTFINDNUMBER")

    Set adoConn = New ADODB.Connection
    With adoConn
        .ConnectionString = adoGetDbConnectString
        .open
    End With
    
    Set adoStream = New ADODB.Stream
    adoStream.open

    'Create a command object to use.
    Set oCommand = New ADODB.Command
    With oCommand
        .CommandType = adCmdStoredProc
        .CommandText = "USP_GetDecisionDetails"
        Set .ActiveConnection = adoConn
        .Parameters.Append .CreateParameter("@p_ApplicationNumber", adVarChar, adParamInput, Len(sApplicationNumber), sApplicationNumber)
        .Parameters.Append .CreateParameter("@p_ApplicationFactFindNumber", adInteger, adParamInput, , CInt(sApplicationFactFindNumber))
        .Properties("Output Stream") = adoStream
        .Properties("XML Root") = "RESPONSE"
        .Execute , , adExecuteStream
        Set .ActiveConnection = Nothing
    End With
        
    adoConn.Close

    IApplicationDO_GetDecisionDetails = adoStream.ReadText()
    
    adoStream.Close

    
GetDecisionDetailsErr:
 
If adoConn.State = adStateOpen Then
       adoConn.Close
    End If
    
    Set adoConn = Nothing
    Set oCommand = Nothing
    Set adoStream = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)
End Function


Private Function IApplicationDO_GetNextApplicationNumber(ByVal vstrChannelID As String) As String
' header ----------------------------------------------------------------------------------
' description:
'   get a new application number, including any prefix
' pass:
'   vstrChannelID   Distribution Channel ID
' return:
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo GetNextApplicationNumberVbErr

    Const strFunctionName As String = "IApplicationDO_GetNextApplicationNumber"


    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
 
    Dim objErrAssist As New ErrAssist
    Dim objXmlElem As IXMLDOMElement
    Dim objXmlTableNode As IXMLDOMNode
    Dim objXmlAssist As New XMLAssist
    Dim objSqlAssist As New SQLAssist
    
    Dim objADOAssist As IADOAssist
    
    Dim strResponse As String
    Dim strCheckDigit As String
    
    If objContext Is Nothing Then
        Set objADOAssist = New ADOAssist
    Else
        Set objADOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    End If
    
    '------------------------------------------------------------------------------------------
    ' get next number
    '------------------------------------------------------------------------------------------
    Dim strNextApplicationNumber As String
    
    Dim adoConnection As ADODB.Connection
    Dim adoCommand As ADODB.Command
    ' DM SYS1946 Sequences.
    #If GENERIC_SQL Then
        Dim adoParameterApplicationNumber As ADODB.Parameter
        Const ProcName = "USP_GETNEXTAPPLICATIONNUMBER"
        Const OutSize = 12
    #Else
        Dim adoRecSet As ADODB.Recordset
    #End If
    
    
'    Dim strSQl As String
'
'    strSQl = "{call sp_NextNumber.GetNextApplicationNumber("
'    strSQl = strSQl & "{resultset 1,"
'    strSQl = strSQl & "NextValue"
'    strSQl = strSQl & "})}"
    
    Set adoConnection = New ADODB.Connection
    With adoConnection
        .ConnectionString = objADOAssist.GetConnStr
        #If GENERIC_SQL Then
        #Else
            .CursorLocation = adUseClient
        #End If
        .open
    End With
    
    Set adoCommand = New ADODB.Command
    With adoCommand
        Set .ActiveConnection = adoConnection
        ' DM SYS1946 Sequences.
        #If GENERIC_SQL Then
            .CommandType = adCmdStoredProc
            .CommandText = ProcName
            ' DM SYS2511 Changed parameter type to adVarChar
            Set adoParameterApplicationNumber = adoCommand.CreateParameter("p_NextNumber", adVarChar, adParamOutput, 12)
            adoCommand.Parameters.Append adoParameterApplicationNumber
        #Else
            .CommandType = adCmdText
            .CommandText = "SELECT SEQ_APPLICATIONNO.NextVal FROM DUAL"
        #End If
        .Prepared = True
    End With
    
    ' Keep getting numbers and calculating the check digit until the
    ' number returned produces a valid Check Digit
    
    'IK_17/05/2006_EP561
    If IsCheckDigitRequired Then
        Do
            #If GENERIC_SQL Then
                adoCommand.Execute
                strNextApplicationNumber = Format(adoParameterApplicationNumber.Value, "0000000")
            #Else
                Set adoRecSet = adoCommand.Execute
                If adoRecSet.EOF Then
                    objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound
                End If
    
                ' APS UNIT TEST REF 84 - Prefix the Next Number with leading zeros
                strNextApplicationNumber = Format(adoRecSet.fields.Item(0).Value, "0000000")
            #End If
            strCheckDigit = GenerateCheckDigit(strNextApplicationNumber)
        Loop Until Not strCheckDigit = vbNullString
    Else
        #If GENERIC_SQL Then
            adoCommand.Execute
            strNextApplicationNumber = Format(adoParameterApplicationNumber.Value, "0000000")
        #Else
            Set adoRecSet = adoCommand.Execute
            If adoRecSet.EOF Then
                objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound
            End If

            ' APS UNIT TEST REF 84 - Prefix the Next Number with leading zeros
            strNextApplicationNumber = Format(adoRecSet.fields.Item(0).Value, "0000000")
        #End If
    End If
    'IK_17/05/2006_EP561
    
    #If GENERIC_SQL Then
        Set adoParameterApplicationNumber = Nothing
    #Else
        adoRecSet.Close
        Set adoRecSet = Nothing
    #End If
    Set adoCommand.ActiveConnection = Nothing
    adoConnection.Close
    Set adoConnection = Nothing
    
    '------------------------------------------------------------------------------------------
    ' add any prefix appropriate to the distribution channel - n.b. can be null
    '------------------------------------------------------------------------------------------
    Dim strCondition As String, strPrefix As String
    Dim blnRecordFound As Boolean
    
    strCondition = " CHANNELID = " & objSqlAssist.FormatString(vstrChannelID)
    
    ' SR 12/06/00 - Do not find prefix, if channelId is empty
    If vstrChannelID <> "" Then
        If IsNull(objADOAssist.GetValueFromTable("DISTRIBUTIONCHANNEL", strCondition, _
                                            "APPLICATIONNUMBERPREFIX", strPrefix, blnRecordFound)) Then
            strPrefix = ""
        End If
        If Not blnRecordFound Then
            objErrAssist.ThrowError TypeName(Me), _
                                    strFunctionName, _
                                    oeMissingPrimaryTag, _
                                    "DISTRIBUTIONCHANNEL : " & vstrChannelID & " not found"
        
        End If
    End If
    
    IApplicationDO_GetNextApplicationNumber = _
        strPrefix & strNextApplicationNumber & strCheckDigit
 
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
        Set objContext = Nothing
    End If
    
    Exit Function

GetNextApplicationNumberVbErr:
   
    #If GENERIC_SQL Then
        Set adoParameterApplicationNumber = Nothing
    #Else
        ' Close the recordset if still open
        If Not adoRecSet Is Nothing Then
            If adoRecSet.State = adStateOpen Then
                adoRecSet.Close
            End If
        End If
        Set adoRecSet = Nothing
    #End If
    
    ' Close the connection if still open
    If Not adoConnection Is Nothing Then
                If adoConnection.State = adStateOpen Then
                        adoConnection.Close
                End If
        End If
            
    Set adoConnection = Nothing
        Set objXmlAssist = Nothing
    
    If objErrAssist.IsSystemError Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set objErrAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    IApplicationDO_GetNextApplicationNumber = ""
    
    ' Re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Private Sub IApplicationDO_CreateApplicationCreditCard(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   Create an instance of the persistant data associated with the table 'ApplicationCreditCard'
'   for each set of data in the request
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo CreateApplicationCreditCardVbErr

    Const strFunctionName As String = "IApplicationDO_CreateApplicationCreditCard"
   
    Dim objErrAssist As New ErrAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
    
    Dim objIDOAssist As IDOAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
                    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If

    Set objIClassDef = New ApplicationClassDef
    
    Set xmlClassDefDoc = objIClassDef.LoadApplicationCreditCardData
    objIDOAssist.Create vxmlTableElement, xmlClassDefDoc

    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Sub

CreateApplicationCreditCardVbErr:
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
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


End Sub

Private Sub IApplicationDO_CreateApplicationFactFind(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   Create an instance of the persistant data associated with the table
'   'ApplicationFactFind' for each set of data in the request
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo CreateApplicationFactFindVbErr

    Const strFunctionName As String = "IApplicationDO_CreateApplicationFactFind"
   
    Dim objErrAssist As New ErrAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
    
    Dim objIDOAssist As IDOAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
                    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadApplicationFactFindData
    objIDOAssist.Create vxmlTableElement, xmlClassDefDoc
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Sub

CreateApplicationFactFindVbErr:
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
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
End Sub

Private Sub IApplicationDO_CreateApplicationPriority(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   APS SYS1993 27/02/01
'   Creates a single instance on the ApplicationPriority table
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo CreateApplicationPriorityVbErr

    Const strFunctionName As String = "IApplicationDO_CreateApplicationPriority"
   
    Dim objErrAssist As New ErrAssist
    Dim lngSequenceNumber As Long
    
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
    
    Dim xmlApplicationPriority As IXMLDOMNode
    Dim xmlStageElem As IXMLDOMElement
    Dim xmlDom As FreeThreadedDOMDocument40
    Dim xmlNode As IXMLDOMNode
    
    Dim objIDOAssist As IDOAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
                    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
        
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadApplicationPriorityData
        
    Set xmlApplicationPriority = vxmlTableElement.selectSingleNode(".//APPLICATIONPRIORITY")
    If xmlApplicationPriority Is Nothing Then
        Set xmlApplicationPriority = vxmlTableElement
    End If
    
    ' Get the next highest sequence number
    lngSequenceNumber = objIDOAssist.GetNextSequenceNumber(vxmlTableElement, xmlClassDefDoc, "APPLICATIONPRIORITY", "APPPRIORITYSEQUENCENO")
      
    Set xmlNode = vxmlTableElement.selectSingleNode(".//APPPRIORITYSEQUENCENO")
    
    If Not xmlNode Is Nothing Then
    
        xmlNode.Text = CStr(lngSequenceNumber)
    Else
           
        Set xmlNode = vxmlTableElement.ownerDocument.createElement("APPPRIORITYSEQUENCENO")
        xmlNode.Text = CStr(lngSequenceNumber)
        xmlApplicationPriority.appendChild xmlNode

    End If
    
    ' Add the current date and time to Request
    Set xmlNode = vxmlTableElement.selectSingleNode(".//DATETIMEASSIGNED")
    
    If xmlNode Is Nothing Then
    
        Set xmlNode = vxmlTableElement.ownerDocument.createElement("DATETIMEASSIGNED")
        xmlApplicationPriority.appendChild xmlNode

    End If
    
    Call AddCurrentDateTimeToRequest("DATETIMEASSIGNED", xmlApplicationPriority)
        
    objIDOAssist.Create vxmlTableElement, xmlClassDefDoc
    
    Set xmlApplicationPriority = Nothing
    Set xmlNode = Nothing
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Sub

CreateApplicationPriorityVbErr:
    
    Set xmlApplicationPriority = Nothing
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
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

End Sub

Private Sub IApplicationDO_CreateApplicationStage(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   CreateApplicationStage an instance of the persistant data associated with this data object
'   for each set of data in the request
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo CreateApplicationStageVbErr

    Const strFunctionName As String = "IApplicationDO_CreateApplicationStage"
   
    Dim objErrAssist As New ErrAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
    Dim xmlCaseSeq As IXMLDOMElement
    Dim strCaseStageSeq As String
    Dim xmlStageElem As IXMLDOMElement
    Dim domReq As FreeThreadedDOMDocument40
    Dim xmlApplicationStage As IXMLDOMNode
    
    Dim objIDOAssist As IDOAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
                    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    ' SYS2030
    Set xmlApplicationStage = vxmlTableElement.selectSingleNode(".//APPLICATIONSTAGE")
    If xmlApplicationStage Is Nothing Then
        Set xmlApplicationStage = vxmlTableElement
    End If
    
    ' Add the current date and time to Request
    Call AddCurrentDateTimeToRequest("DATETIME", xmlApplicationStage)
    
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadApplicationStageData
    
    Set domReq = New FreeThreadedDOMDocument40
    domReq.validateOnParse = False
    domReq.setProperty "NewParser", True
    domReq.appendChild xmlApplicationStage.cloneNode(True)
    
    Set xmlStageElem = m_objXmlAssist.GetNode(domReq.documentElement, "//APPLICATIONSTAGE")
    
    ' Need to check the CASESTAGESEQUENCENO value
    strCaseStageSeq = m_objXmlAssist.GetNodeValue(xmlStageElem, ".//CASESTAGESEQUENCENO", False)
    
    Set xmlCaseSeq = xmlStageElem.appendChild(xmlStageElem.ownerDocument.createElement("CASESEQUENCENO"))
    
    If Len(strCaseStageSeq) = 0 Then
        strCaseStageSeq = "1"
    End If
    
    xmlCaseSeq.Text = strCaseStageSeq

    objIDOAssist.Create domReq.documentElement, xmlClassDefDoc
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlApplicationStage = Nothing
    Set domReq = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Sub

CreateApplicationStageVbErr:
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set xmlApplicationStage = Nothing
    Set domReq = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
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

End Sub

Private Sub IApplicationDO_CreateApplicationVerification(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   Create an instance of the persistant data associated with the table 'ApplicationVerification'
'   for each set of data in the request
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo CreateApplicationVerificationVbErr

    Const strFunctionName As String = "IApplicationDO_CreateApplicationVerification"
   
    Dim objErrAssist As New ErrAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
    
    Dim objIDOAssist As IDOAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
                    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If

    Set objIClassDef = New ApplicationClassDef
    
    ' APPLICATIONVERIFICATION:
    Set xmlClassDefDoc = objIClassDef.LoadApplicationVerificationData
    objIDOAssist.Create vxmlTableElement, xmlClassDefDoc

    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Sub

CreateApplicationVerificationVbErr:
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
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

End Sub

Private Sub IApplicationDO_CreateApplicationUnderwriting(ByVal vxmlTableElement As IXMLDOMElement)

On Error GoTo CreateApplicationUnderwritingVbErr

    Const strFunctionName As String = "IApplicationDO_CreateApplicationUnderwriting"
   
    Dim objErrAssist As New ErrAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
    
    Dim objIDOAssist As IDOAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
                    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If

    Set objIClassDef = New ApplicationClassDef
    
    ' APPLICATIONVERIFICATION:
    Set xmlClassDefDoc = objIClassDef.LoadApplicationUnderwritingData
    objIDOAssist.Create vxmlTableElement, xmlClassDefDoc

    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Sub

CreateApplicationUnderwritingVbErr:
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
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

End Sub

Private Sub IApplicationDO_CreateCustomerRelationship(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   CreateCustomerRelationship an instance of the persistant data associated with this data object
'   for each set of data in the request
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo CreateCustomerRelationshipVbErr

    Const strFunctionName As String = "IApplicationDO_CreateCustomerRelationship"
   
    Dim objErrAssist As New ErrAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
    
    Dim objIDOAssist As IDOAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
                    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadCustomerRelationshipData()
    objIDOAssist.Create vxmlTableElement, xmlClassDefDoc
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Sub

CreateCustomerRelationshipVbErr:
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
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
    
End Sub

Private Sub IApplicationDO_CreateFinancialSummary(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   CreateFinancialSummary an instance of the persistant data associated with this data object
'   for each set of data in the request
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo CreateFinancialSummaryVbErr

    Const strFunctionName As String = "IApplicationDO_CreateFinancialSummary"
   
    Dim objErrAssist As New ErrAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
    
    Dim objIDOAssist As IDOAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
                    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadFinancialSummaryData()
    objIDOAssist.Create vxmlTableElement, xmlClassDefDoc
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Sub

CreateFinancialSummaryVbErr:
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
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
    
End Sub

Private Sub IApplicationDO_CreateLoanProperty(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   Create an instance of the persistant data associated with the table 'MortgageClaculatorDetails'
'   for each set of data in the request
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo CreateLoanPropertyVbErr

    Const strFunctionName As String = "IApplicationDO_CreateLoanProperty"
   
    Dim objErrAssist As New ErrAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
    
    Dim objIDOAssist As IDOAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
                    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If

    Set objIClassDef = New ApplicationClassDef
    
    ' NEWPROPERTY
    Set xmlClassDefDoc = objIClassDef.LoadNewPropertyData
    objIDOAssist.Create vxmlTableElement, xmlClassDefDoc

    ' NEWLOAN
    Set xmlClassDefDoc = objIClassDef.LoadNewLoanData
    objIDOAssist.Create vxmlTableElement, xmlClassDefDoc

    ' APPLICATIONFACTFIND
    Set xmlClassDefDoc = objIClassDef.LoadApplicationFactFindData
    objIDOAssist.Update vxmlTableElement, xmlClassDefDoc

    ' SHAREDOWNERSHIPDETAILS
    Set xmlClassDefDoc = objIClassDef.LoadSharedOwnershipDetailsData()
    objIDOAssist.Create vxmlTableElement, xmlClassDefDoc
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Sub

CreateLoanPropertyVbErr:
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
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
End Sub


Private Sub IApplicationDO_CreatePropertyDetails(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   Creates NEWPROPERTY & NEWLOAN records associated with the APLICATIONFACTFIND
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo CreatePropertyDetailsVbErr

    Const strFunctionName As String = "IApplicationDO_CreatePropertyDetails"
       
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
    Dim bNewPropertySummary As Boolean
    Dim objIDOAssist As IDOAssist
    Dim objINewPropertyDO As INewPropertyDO
    Dim xmlSharedOwnershipDetails As IXMLDOMNode
    Dim xmlLeaseholdNode As IXMLDOMNode
    Dim xmlNewLoan As IXMLDOMNode
   
    '~ MF MAR20 If NEWPROPERTYSUMMARY tag is present then save RoomType data also
    bNewPropertySummary = Not vxmlTableElement.getElementsByTagName("NEWPROPERTYSUMMARY") Is Nothing

    If m_objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
        If bNewPropertySummary Then
            Set objINewPropertyDO = New NewPropertyDO
        End If
    Else
        Set objIDOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
        If bNewPropertySummary Then
            Set objINewPropertyDO = m_objContext.CreateInstance(App.Title & ".NewPropertyDO")
        End If
    End If

    Set objIClassDef = New ApplicationClassDef
    
    ' NEWPROPERTY
    Set xmlClassDefDoc = objIClassDef.LoadNewPropertyData
    objIDOAssist.Create vxmlTableElement, xmlClassDefDoc
    
    ' NEWPROPERTYLEASEHOLD    JD MAR40 added
    Set xmlClassDefDoc = objIClassDef.LoadNewPropertyLeaseholdData
    'MAR109 GHun NEWPROPERTYLEASEHOLD is not mandatory
    Set xmlLeaseholdNode = m_objXmlAssist.GetNode(vxmlTableElement, "//NEWPROPERTY/NEWPROPERTYLEASEHOLD", False)
    If Not xmlLeaseholdNode Is Nothing Then
        objIDOAssist.Create xmlLeaseholdNode, xmlClassDefDoc
    End If

    ' NEWLOAN
    Set xmlClassDefDoc = objIClassDef.LoadNewLoanData
    'MAR109 GHun NEWLOAN is not mandatory
    Set xmlNewLoan = m_objXmlAssist.GetNode(vxmlTableElement, "//NEWPROPERTY/NEWLOAN", False)
    If Not xmlNewLoan Is Nothing Then
        objIDOAssist.Create vxmlTableElement, xmlClassDefDoc
    End If

    '~ NEWPROPERTYROOMTYPE
    If bNewPropertySummary Then
        objINewPropertyDO.UpdateNewPropertyRoomDetails vxmlTableElement
    End If

    ' APPLICATIONFACTFIND
    Set xmlClassDefDoc = objIClassDef.LoadApplicationFactFindData
    objIDOAssist.Update vxmlTableElement, xmlClassDefDoc

    ' SHAREDOWNERSHIPDETAILS
    Set xmlClassDefDoc = objIClassDef.LoadSharedOwnershipDetailsData()
    'MAR109 GHun SHAREDOWNERSHIPDETAILS is not mandatory
    Set xmlSharedOwnershipDetails = m_objXmlAssist.GetNode(vxmlTableElement, "//NEWPROPERTY/SHAREDOWNERSHIPDETAILS", False)
    If Not xmlSharedOwnershipDetails Is Nothing Then
        objIDOAssist.Create vxmlTableElement, xmlClassDefDoc
    End If
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objINewPropertyDO = Nothing
    
    Set xmlSharedOwnershipDetails = Nothing
    Set xmlLeaseholdNode = Nothing
    Set xmlNewLoan = Nothing
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

    Exit Sub

CreatePropertyDetailsVbErr:
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set xmlSharedOwnershipDetails = Nothing
    Set xmlLeaseholdNode = Nothing
    Set xmlNewLoan = Nothing
    
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


Private Sub IApplicationDO_CreateMenuItemStatus(ByVal vxmlTableElement As IXMLDOMElement)

On Error GoTo CreateMenuItemStatusVbErr

    Const strFunctionName As String = "IApplicationDO_CreateMenuItemStatus"
   
    Dim objErrAssist As New ErrAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
    
    Dim objIDOAssist As IDOAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
                    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadMenuItemStatusData
    objIDOAssist.Create vxmlTableElement, xmlClassDefDoc
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Sub

CreateMenuItemStatusVbErr:
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
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
End Sub

Private Sub IApplicationDO_CreateMortgageCalcDetails(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   Create an instance of the persistant data associated with the table 'MortgageClaculatorDetails'
'   for each set of data in the request
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo CreateMortgageCalcDetailsVbErr

    Const strFunctionName As String = "IApplicationDO_CreateMortgageCalcDetails"
   
    Dim objErrAssist As New ErrAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
    
    Dim objIDOAssist As IDOAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
                    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadMortgageCalcDetailsData
    objIDOAssist.Create vxmlTableElement, xmlClassDefDoc
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Sub

CreateMortgageCalcDetailsVbErr:
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
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
End Sub

Private Sub IApplicationDO_CreateNewLoan(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   Create an instance of the persistant data associated with the table 'NewLoan'
'   for each set of data in the request
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo CreateNewLoanVbErr

    Const strFunctionName As String = "IApplicationDO_CreateNewLoan"
   
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
    
    Dim objIDOAssist As IDOAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
                    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadNewLoanData
    objIDOAssist.Create vxmlTableElement, xmlClassDefDoc
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Sub

CreateNewLoanVbErr:
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
        
    Dim objErrAssist As New ErrAssist
        
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
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

End Sub

Private Sub IApplicationDO_CreateOtherInsuranceCompany(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   Create an instance of the persistant data associated with this data object
'   for each set of data in the request
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo CreateOtherInsuranceCompanyVbErr

    Const strFunctionName As String = "IApplicationDO_CreateOtherInsuranceCompany"
   
    Dim objErrAssist As New ErrAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
    
    Dim objIDOAssist As IDOAssist
    Dim objXmlAssist As New XMLAssist

    Dim xmlNode        As IXMLDOMNode
    Dim xmlReturnXML As IXMLDOMNode
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
                    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If

    ' Save Other Insurance Company
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadOtherInsuranceCompanyData
    objIDOAssist.Create vxmlTableElement, xmlClassDefDoc
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlNode = Nothing
    Set xmlReturnXML = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Sub

CreateOtherInsuranceCompanyVbErr:
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlNode = Nothing
    Set xmlReturnXML = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
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
    

End Sub

Private Sub IApplicationDO_CreateOtherResident(ByVal vxmlTableElement As IXMLDOMElement, _
                                               ByRef rxmlKeysNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   CreateOtherResident an instance of the persistant data associated with this data object
'   for each set of data in the request
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo CreateOtherResidentVbErr

    Const strFunctionName As String = "IApplicationDO_CreateOtherResident"
   
    Dim objErrAssist As New ErrAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
    
    Dim objIDOAssist As IDOAssist
    Dim objXmlAssist As New XMLAssist
    Dim objPersonDO As IPersonDO

    Dim xmlPersonNode        As IXMLDOMNode, _
        xmlOtherResidentNode As IXMLDOMNode
    Dim xmlPersonGUIDNode As IXMLDOMNode
    Dim xmlPersonReturnXML As IXMLDOMNode
    Dim strPersonGUID As String
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
                    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
        Set objPersonDO = New PersonDO
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
        Set objPersonDO = objContext.CreateInstance(App.Title & ".PersonDO")
    End If

    ' Get PERSON and OTHERRESIDENT nodes from the XML
    Set xmlPersonNode = vxmlTableElement.selectSingleNode("PERSON")
    If xmlPersonNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingPrimaryTag, "PERSON tag not found"
    End If

    If vxmlTableElement.nodeName = "OTHERRESIDENT" Then
        Set xmlOtherResidentNode = vxmlTableElement
    Else
        Set xmlOtherResidentNode = vxmlTableElement.selectSingleNode("OTHERRESIDENT")
    End If

    If xmlOtherResidentNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingPrimaryTag, "OTHERRESIDENT tag not found"
    End If

    '
    ' Save Person
    '
    strPersonGUID = objXmlAssist.GetTagValue(xmlPersonNode, "PERSONGUID")

    Set xmlPersonReturnXML = objPersonDO.Save(xmlPersonNode)
    If strPersonGUID = "" Then
        strPersonGUID = objXmlAssist.GetTagValue(xmlPersonReturnXML, "PERSONGUID")
    End If

    '
    ' Save Other Resident
    '
    ' Insert person GUID into other resident XML
    Set xmlPersonGUIDNode = vxmlTableElement.selectSingleNode("PERSONGUID")
    If xmlPersonGUIDNode Is Nothing Then
        ' No GUID tag was found so create it now
        Set xmlPersonGUIDNode = xmlOtherResidentNode.appendChild( _
                                    vxmlTableElement.ownerDocument.createElement("PERSONGUID"))
    End If
    xmlPersonGUIDNode.Text = strPersonGUID

    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadOtherResidentData
    objIDOAssist.GenerateSequenceNumber vxmlTableElement, xmlClassDefDoc, "OTHERRESIDENTSEQUENCENUMBER"
    objIDOAssist.Create vxmlTableElement, xmlClassDefDoc
    
    ' Add person GUID to the response XML
    rxmlKeysNode.appendChild(rxmlKeysNode.ownerDocument.createElement("PERSONGUID")).Text = strPersonGUID
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objPersonDO = Nothing
    Set xmlPersonNode = Nothing
    Set xmlOtherResidentNode = Nothing
    Set xmlPersonGUIDNode = Nothing
    Set xmlPersonReturnXML = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Sub

CreateOtherResidentVbErr:
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objXmlAssist = Nothing
    Set objPersonDO = Nothing
    Set xmlPersonNode = Nothing
    Set xmlOtherResidentNode = Nothing
    Set xmlPersonGUIDNode = Nothing
    Set xmlPersonReturnXML = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
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
    
End Sub

Private Sub IApplicationDO_CreateSharedOwnershipDetails(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   Create an instance of the persistant data associated with the table 'MortgageClaculatorDetails'
'   for each set of data in the request
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo CreateSharedOwnershipDetailsVbErr

    Const strFunctionName As String = "IApplicationDO_CreateSharedOwnershipDetails"
   
    Dim objErrAssist As New ErrAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
    
    Dim objIDOAssist As IDOAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
                    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadSharedOwnershipDetailsData
    objIDOAssist.Create vxmlTableElement, xmlClassDefDoc
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Sub

CreateSharedOwnershipDetailsVbErr:
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
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
End Sub

Private Sub IApplicationDO_CreateUserHistory(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   CreateUserHistory an instance of the persistant data associated with this data object
'   for each set of data in the request
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo CreateUserHistoryVbErr

    Const strFunctionName As String = "IApplicationDO_CreateUserHistory"
   
    Dim objErrAssist As New ErrAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
    Dim objXmlAssist As XMLAssist
    
    Dim objIDOAssist As IDOAssist
    Dim objContext As ObjectContext
    
    Set objContext = GetObjectContext()
    Set objXmlAssist = New XMLAssist
    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    ' DJP, if the date isn't passed in, create it
    Dim strVal As String
    Dim xmlUserHistory As IXMLDOMElement
    Dim xmlDom As FreeThreadedDOMDocument40
    Const strDateNode = "USERHISTORYDATE"
    
    ' Due to problems with other methods calling this and the owner document's document element being
    ' different to our vxmlTableElement.xml, GetNode fails. So, create a new domdocument and copy
    ' everything into it, then use the document element. Because at this point USERHISTORY is the root
    ' node, we have to start at the top, and if we start at the top, we get a totally different bit
    ' of XML unless we create a new dom document

    Set xmlDom = New FreeThreadedDOMDocument40
    xmlDom.validateOnParse = False
    xmlDom.setProperty "NewParser", True
    Set xmlDom.documentElement = vxmlTableElement.cloneNode(True)
    
    Set xmlUserHistory = objXmlAssist.GetNode(xmlDom.documentElement, "//USERHISTORY", True)
    'Set xmlUserHistory = objXmlAssist.GetNode(vxmlTableElement, ".//USERHISTORY", False)
    strVal = objXmlAssist.GetNodeValue(xmlUserHistory, "./" & strDateNode, False)
    
    If Len(strVal) = 0 Then
        Set xmlDom = xmlUserHistory.ownerDocument
        xmlUserHistory.appendChild(xmlDom.createElement(strDateNode)).Text = Now
    
    End If
    
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoaduserHistoryData
    objIDOAssist.Create xmlDom.documentElement, xmlClassDefDoc
    Set xmlDom = Nothing
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Sub

CreateUserHistoryVbErr:
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
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

End Sub

Private Sub IApplicationDO_CreateVerification(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   CreateVerification an instance of the persistant data associated with this data object
'   for each set of data in the request
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo CreateVerificationVbErr

    Const strFunctionName As String = "IApplicationDO_CreateVerification"
   
    Dim objErrAssist As New ErrAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
    
    Dim objIDOAssist As IDOAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
                    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadVerificationData
    objIDOAssist.GenerateSequenceNumber vxmlTableElement, xmlClassDefDoc, "VERIFICATIONSEQUENCENUMBER"
    objIDOAssist.Create vxmlTableElement, xmlClassDefDoc
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Sub

CreateVerificationVbErr:
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
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
    
End Sub

Private Sub IApplicationDO_DeleteApplicationStage(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   Deletes a single instance of the persistant data associated with the table 'ApplicationStage'
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo DeleteApplicationStageVbErr

    Const strFunctionName As String = "IApplicationDO_DeleteApplicationStage"

    Dim objErrAssist As New ErrAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
            
    Dim objIDOAssist As IDOAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
            
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadApplicationStageData()
    objIDOAssist.Delete vxmlTableElement, xmlClassDefDoc
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Sub

DeleteApplicationStageVbErr:

    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
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
End Sub

Private Sub IApplicationDO_DeleteDependantForCustomer(ByVal vxmlTableElement As IXMLDOMElement, Optional ByVal bRaiseErr As Boolean = True)
' header ----------------------------------------------------------------------------------
' description:
'   DeleteDependantForCustomer a single instance of the persistant data associated with this
'   data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo DeleteDependantForCustomerVbErr

    Const strFunctionName As String = "IApplicationDO_DeleteDependantForCustomer"

    Dim objErrAssist As New ErrAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
            
    Dim objIADOAssist As IADOAssist, _
        objSqlAssist As New SQLAssist, _
        objXmlAssist As New XMLAssist

    Dim strCustomerNumber            As String, _
        strCustomerVersionNumber     As String, _
        strDependantSequenceNumber   As String
    Dim cmd As ADODB.Command
    
    If objContext Is Nothing Then
        Set objIADOAssist = New ADOAssist
    Else
        Set objIADOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    End If
            
    strCustomerNumber = objXmlAssist.GetTagValue(vxmlTableElement, "CUSTOMERNUMBER")
    strCustomerVersionNumber = objXmlAssist.GetTagValue(vxmlTableElement, "CUSTOMERVERSIONNUMBER")
    strDependantSequenceNumber = objXmlAssist.GetTagValue(vxmlTableElement, "DEPENDANTSEQUENCENUMBER")

    If (strCustomerNumber = "") Or _
       (strCustomerVersionNumber = "") Or _
       (strDependantSequenceNumber = "") Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingPrimaryTag
    End If


    Set cmd = New ADODB.Command
    With cmd
        .ActiveConnection = objIADOAssist.GetConnStr
        .CommandType = adCmdStoredProc
        .CommandText = "DELETEDEPENDANTFORCUSTOMER"
    'SYS1730 - Use the single stored proc calling convetion. Do not use Parameters.Refresh().
    '#If GENERIC_SQL Then
        Dim prmOtherResidentCount As ADODB.Parameter
        .Parameters.Append .CreateParameter("strCUSTOMERNUMBER", adVarChar, adParamInput, Len(strCustomerNumber), strCustomerNumber)
        .Parameters.Append .CreateParameter("intCUSTOMERVERSIONNUMBER", adInteger, adParamInput, , strCustomerVersionNumber)
        .Parameters.Append .CreateParameter("intDEPENDANTSEQUENCENUMBER", adInteger, adParamInput, , strDependantSequenceNumber)
        Set prmOtherResidentCount = .CreateParameter("intOtherResidentCount", adInteger, adParamOutput)
        .Parameters.Append prmOtherResidentCount
        .Execute , , adCmdStoredProc
        If prmOtherResidentCount.Value > 0 And bRaiseErr Then
            ' Dependant was an other resident but all other resident records have now been deleted - warn the user
            objErrAssist.ThrowError TypeName(Me), strFunctionName, 553
        End If
    '#Else
    '        .Parameters.Refresh
    '        .Parameters("strCUSTOMERNUMBER").Value = strCustomerNumber
    '        .Parameters("intCUSTOMERVERSIONNUMBER").Value = strCustomerVersionNumber
    '        .Parameters("intDEPENDANTSEQUENCENUMBER").Value = strDependantSequenceNumber
    '        .Execute , , adCmdStoredProc
    '        If .Parameters("intOtherResidentCount").Value > 0 And bRaiseErr Then
    '            ' Dependant was an other resident but all other resident records have now been deleted - warn the user
    '            objErrAssist.ThrowError TypeName(Me), strFunctionName, 553
    '        End If
    '#End If
    'SYS1730 - End.
    End With

    Set objIADOAssist = Nothing
    Set objErrAssist = Nothing
    Set cmd = Nothing
    Set objSqlAssist = Nothing
    Set objXmlAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Sub

DeleteDependantForCustomerVbErr:

    Set objIADOAssist = Nothing
    Set cmd = Nothing
    Set objSqlAssist = Nothing
    Set objXmlAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
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
End Sub

Private Sub IApplicationDO_DeleteMortgageCalcDetails(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   Deletes a single instance of the persistant data associated with the table 'MortgageCalcDetails'
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo DeleteMortgageCalcDetailsVbErr

    Const strFunctionName As String = "IApplicationDO_DeleteMortgageCalcDetails"

    Dim objErrAssist As New ErrAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
            
    Dim objIDOAssist As IDOAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
            
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadMortgageCalcDetailsData()
    objIDOAssist.Delete vxmlTableElement, xmlClassDefDoc
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Sub

DeleteMortgageCalcDetailsVbErr:

    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
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
End Sub

Private Sub IApplicationDO_DeleteOtherResident(ByVal vxmlTableElement As IXMLDOMElement)
    
    On Error GoTo DeleteOtherResidentVbErr

    Const strFunctionName As String = "IApplicationDO_DeleteOtherResident"

    Dim objErrAssist As New ErrAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
            
    Dim objIADOAssist As IADOAssist, _
        objSqlAssist As New SQLAssist, _
        objXmlAssist As New XMLAssist

    Dim strApplicationNumber           As String, _
        strApplicationFactFindNumber   As String, _
        strOtherResidentSequenceNumber As String
    Dim cmd As ADODB.Command
    
    If objContext Is Nothing Then
        Set objIADOAssist = New ADOAssist
    Else
        Set objIADOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    End If
            
    strApplicationNumber = objXmlAssist.GetTagValue(vxmlTableElement, "APPLICATIONNUMBER")
    strApplicationFactFindNumber = objXmlAssist.GetTagValue(vxmlTableElement, "APPLICATIONFACTFINDNUMBER")
    strOtherResidentSequenceNumber = objXmlAssist.GetTagValue(vxmlTableElement, "OTHERRESIDENTSEQUENCENUMBER")

    If (strApplicationNumber = "") Or _
       (strApplicationFactFindNumber = "") Or _
       (strOtherResidentSequenceNumber = "") Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingPrimaryTag
    End If


    Set cmd = New ADODB.Command
    With cmd
        .ActiveConnection = objIADOAssist.GetConnStr
        .CommandType = adCmdStoredProc
        .CommandText = "DELETEOTHERRESIDENT"
        'ASt 12/03/01 CC012 Now compatible with SQL Server and Oracle
    'SYS1730 - Use the single stored proc calling convetion. Do not use Parameters.Refresh().
    '#If GENERIC_SQL Then
            Dim prmDependantCount As ADODB.Parameter
            .Parameters.Append .CreateParameter("strAPPLICATIONNUMBER", adVarChar, adParamInput, Len(strApplicationNumber), strApplicationNumber)
            .Parameters.Append .CreateParameter("intAPPLICATIONFACTFINDNUMBER", adInteger, adParamInput, , strApplicationFactFindNumber)
            .Parameters.Append .CreateParameter("intOTHERRESIDENTSEQUENCENUMBER", adInteger, adParamInput, , strOtherResidentSequenceNumber)
            Set prmDependantCount = .CreateParameter("intDependantCount", adInteger, adParamOutput)
            .Parameters.Append prmDependantCount
            .Execute , , adCmdStoredProc
            If prmDependantCount.Value > 0 Then
                ' Other resident was deleted but the person was not as they still exist as a dependant - warn the user
                objErrAssist.ThrowError TypeName(Me), strFunctionName, 554
            End If
    '#Else
    '        .Parameters.Refresh
    '        .Parameters("strAPPLICATIONNUMBER").Value = strApplicationNumber
    '        .Parameters("intAPPLICATIONFACTFINDNUMBER").Value = strApplicationFactFindNumber
    '        .Parameters("intOTHERRESIDENTSEQUENCENUMBER").Value = strOtherResidentSequenceNumber
    '        .Execute , , adCmdStoredProc
    '        If .Parameters("intDependantCount").Value > 0 Then
    '            ' Other resident was deleted but the person was not as they still exist as a dependant - warn the user
    '            objErrAssist.ThrowError TypeName(Me), strFunctionName, 554
    '        End If
    '#End If
    'SYS1730 - End.
    End With
    
    Set objIADOAssist = Nothing
    Set objErrAssist = Nothing
    Set cmd = Nothing
    Set objSqlAssist = Nothing
    Set objXmlAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Sub

DeleteOtherResidentVbErr:

    Set objIADOAssist = Nothing
    Set cmd = Nothing
    Set objSqlAssist = Nothing
    Set objXmlAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
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
End Sub

Private Sub IApplicationDO_DeleteUserHistory(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   DeleteUserHistory a single instance of the persistant data associated with this
'   data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo DeleteUserHistoryVbErr

    Const strFunctionName As String = "IApplicationDO_DeleteUserHistory"

    Dim objErrAssist As New ErrAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
            
    Dim objIDOAssist As IDOAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
            
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoaduserHistoryData()
    objIDOAssist.Delete vxmlTableElement, xmlClassDefDoc
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Sub

DeleteUserHistoryVbErr:

    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
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
End Sub

Private Sub IApplicationDO_DeleteVerification(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   DeleteVerification a single instance of the persistant data associated with this
'   data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo DeleteVerificationVbErr

    Const strFunctionName As String = "IApplicationDO_DeleteVerification"

    Dim objErrAssist As New ErrAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
            
    Dim objIDOAssist As IDOAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
            
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadVerificationData()
    objIDOAssist.Delete vxmlTableElement, xmlClassDefDoc
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Sub

DeleteVerificationVbErr:

    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
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
End Sub

Private Function IApplicationDO_FindApplicationFactFindList( _
                        ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode

' header ----------------------------------------------------------------------------------
' description:
'   Get the data for all instances of the persistant data associated with
'   the table 'ApplicationFactFind' for the values supplied
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo FindApplicationFactFindListVbErr

    Const strFunctionName As String = "IApplicationDO_FindApplicationFactFindList"

    Dim objErrAssist As New ErrAssist
    
    Dim objIDOAssist As IDOAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
         
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
     
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadApplicationFactFindData()
    Set IApplicationDO_FindApplicationFactFindList = objIDOAssist.FindListMultiple(vxmlTableElement, xmlClassDefDoc)

    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objErrAssist = Nothing
    Set objIDOAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

FindApplicationFactFindListVbErr:
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
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

Private Function IApplicationDO_FindApplicationStageList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for all instances of the persistant data associated with
'   the table 'ApplicationStage' for the values supplied
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo FindApplicationStageListVbErr

    Const strFunctionName As String = "IApplicationDO_FindApplicationStageList"

    Dim objErrAssist As New ErrAssist
    
    Dim objIDOAssist As IDOAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
         
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
     
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadApplicationStageData()
    Set IApplicationDO_FindApplicationStageList = objIDOAssist.FindListMultiple(vxmlTableElement, xmlClassDefDoc)

    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objErrAssist = Nothing
    Set objIDOAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

FindApplicationStageListVbErr:
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
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

Private Function IApplicationDO_FindCustomerAddressListAndSalutation(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for all instances of the persistant data associated with
'   this data object for the values supplied
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo FindCustomerAddressListAndSalutationVbErr

    Const strFunctionName As String = "IApplicationDO_FindCustomerAddressListAndSalutation"

    Dim objErrAssist As New ErrAssist
    
    Dim objIDOAssist As IDOAssist, _
        objXmlAssist As New XMLAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef

    Dim xmlApplicationNode     As IXMLDOMNode, _
        xmlCustomerAddressNode As IXMLDOMNode
    Dim xmlApplicationResponse     As IXMLDOMNode, _
        xmlCustomerAddressResponse As IXMLDOMNode
    Dim objICustomerBO As ICustomerBO
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
         
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
        Set objICustomerBO = New CustomerBO
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
        Set objICustomerBO = objContext.CreateInstance(gstrCUSTOMER_COMPONENT & ".CustomerBO")
    End If

    ' Find the APPLICATION criteria and CUSTOMERADDRESS criteria
    Set xmlApplicationNode = vxmlTableElement

    Set xmlCustomerAddressNode = vxmlTableElement.selectSingleNode(".//CUSTOMERADDRESSLIST")
    If xmlCustomerAddressNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingPrimaryTag, "CUSTOMERADDRESSLIST" & " tag not found"
    End If

    ' Fetch the address salutation from APPLICATION
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadApplicationData
    Set xmlApplicationResponse = objIDOAssist.GetData(xmlApplicationNode, xmlClassDefDoc)

    ' Fetch customer address(es)
    On Error GoTo RecordNotFoundErr
    Set xmlCustomerAddressResponse = objICustomerBO.FindCustomerAddressList(xmlCustomerAddressNode)
    On Error GoTo FindCustomerAddressListAndSalutationVbErr
    xmlApplicationResponse.appendChild xmlCustomerAddressResponse

    Set IApplicationDO_FindCustomerAddressListAndSalutation = xmlApplicationResponse

    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlApplicationNode = Nothing
    Set xmlCustomerAddressNode = Nothing
    Set xmlApplicationResponse = Nothing
    Set xmlCustomerAddressResponse = Nothing
    Set objICustomerBO = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

RecordNotFoundErr:
    If objErrAssist.GetOmigaErrorNumber(Err.Number) = oeRecordNotFound Then
        ' No customer addresses found - still want to return the address salutation though
        Set xmlCustomerAddressResponse = vxmlTableElement.ownerDocument.createElement("CUSTOMERADDRESSLIST")
        Resume Next
    End If

FindCustomerAddressListAndSalutationVbErr:
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlApplicationNode = Nothing
    Set xmlCustomerAddressNode = Nothing
    Set xmlApplicationResponse = Nothing
    Set xmlCustomerAddressResponse = Nothing
    Set objICustomerBO = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
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

Private Function IApplicationDO_FindDependantsForCustomer(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for all instances of the persistant data associated with
'   this data object for the values supplied
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo FindDependantsForCustomerVbErr

    Const strFunctionName As String = "IApplicationDO_FindDependantsForCustomer"

    Dim objErrAssist As New ErrAssist
    
    Dim objIDOAssist As IDOAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef

    Dim objXmlAssist As New XMLAssist
    Dim xmlReturn As IXMLDOMNode
    Dim xmlORRequest  As IXMLDOMNode, _
        xmlORResponse As IXMLDOMNode
    Dim strApplicationNumber         As String, _
        strApplicationFactFindNumber As String
    Dim xmlDependantNodeList As IXMLDOMNodeList, _
        xmlDependant         As IXMLDOMNode
    Dim xmlORNodeList As IXMLDOMNodeList, _
        xmlOR         As IXMLDOMNode
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
         
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If

    '
    ' Retrieve the dependant + person data
    '
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadDependantDetailsData()
    Set xmlReturn = objIDOAssist.FindListMultiple(vxmlTableElement, _
            xmlClassDefDoc, vstritemname:="DEPENDANT")

    Set xmlDependantNodeList = xmlReturn.selectNodes(".//DEPENDANTDETAILS/DEPENDANT")
    For Each xmlDependant In xmlDependantNodeList
        objXmlAssist.PromoteNode xmlReturn, xmlDependant
    Next xmlDependant

    '
    ' Retrieve the other resident data
    '
    strApplicationNumber = objXmlAssist.GetTagValue(vxmlTableElement, "APPLICATIONNUMBER")
    strApplicationFactFindNumber = objXmlAssist.GetTagValue(vxmlTableElement, "APPLICATIONFACTFINDNUMBER")

    If (strApplicationNumber = "") Or (strApplicationFactFindNumber = "") Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingPrimaryTag
    End If

    Set xmlORRequest = vxmlTableElement.ownerDocument.createElement("OTHERRESIDENT")
    xmlORRequest.appendChild(vxmlTableElement.ownerDocument.createElement("APPLICATIONNUMBER")).Text _
        = strApplicationNumber
    xmlORRequest.appendChild(vxmlTableElement.ownerDocument.createElement("APPLICATIONFACTFINDNUMBER")).Text _
        = strApplicationFactFindNumber
    On Error GoTo RecordNotFoundErr
    Set xmlORResponse = IApplicationDO_FindOtherResidentList(xmlORRequest)
    On Error GoTo FindDependantsForCustomerVbErr

    '
    ' Append the other resident data to the dependant data as appropriate
    '
    If Not (xmlORResponse Is Nothing) Then
        Set xmlORNodeList = xmlORResponse.selectNodes("OTHERRESIDENT")
        Set xmlDependantNodeList = xmlReturn.selectNodes("DEPENDANT")
        For Each xmlOR In xmlORNodeList
            For Each xmlDependant In xmlDependantNodeList
                If objXmlAssist.GetTagValue(xmlOR, "PERSONGUID") = objXmlAssist.GetTagValue(xmlDependant, "PERSONGUID") Then
                    ' Other resident 'record' corresponds to this dependant record; so append
                    xmlDependant.appendChild xmlOR
                End If
            Next xmlDependant
        Next xmlOR
    End If

    Set IApplicationDO_FindDependantsForCustomer = xmlReturn

    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objErrAssist = Nothing
    Set objIDOAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlReturn = Nothing
    Set xmlDependant = Nothing
    Set xmlDependantNodeList = Nothing
    Set xmlORRequest = Nothing
    Set xmlORResponse = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

RecordNotFoundErr:
    If objErrAssist.GetOmigaErrorNumber(Err.Number) = oeRecordNotFound Then
        Resume Next
    End If

FindDependantsForCustomerVbErr:
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlReturn = Nothing
    Set xmlDependant = Nothing
    Set xmlDependantNodeList = Nothing
    Set xmlORRequest = Nothing
    Set xmlORResponse = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
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

Private Function IApplicationDO_FindMainApplicants(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for all instances of the persistant data associated with
'   this data object for the values supplied
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo FindMainApplicantsVbErr

    Const strFunctionName As String = "IApplicationDO_FindMainApplicants"

    Dim objErrAssist As New ErrAssist
    
    Dim objIDOAssist As IDOAssist, _
        objIADOAssist As IADOAssist, _
        objXmlAssist As New XMLAssist, _
        objSqlAssist As New SQLAssist

    Dim objIClassDef As IApplicationClassDef
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim strApplicationNumber         As String, _
        strApplicationFactFindNumber As String
    Dim strSQL As String
    Dim cmd As ADODB.Command, _
        rst As ADODB.Recordset
    Dim xmlDoc As New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
         
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
        Set objIADOAssist = New ADOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
        Set objIADOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    End If

    ' Get application number from XML
    strApplicationNumber = objXmlAssist.GetTagValue(vxmlTableElement, "APPLICATIONNUMBER")
    strApplicationFactFindNumber = objXmlAssist.GetTagValue(vxmlTableElement, "APPLICATIONFACTFINDNUMBER")

    If (strApplicationNumber = "") Or (strApplicationFactFindNumber = "") Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingPrimaryTag
    End If

    xmlDoc.appendChild xmlDoc.createElement("MAINAPPLICANTLIST")

    strSQL = "SELECT * FROM CUSTOMERROLE WHERE " & _
            " APPLICATIONNUMBER = " & objSqlAssist.FormatString(strApplicationNumber) & _
            " AND APPLICATIONFACTFINDNUMBER = " & strApplicationFactFindNumber & _
            " AND CUSTOMERROLETYPE = 1 AND CUSTOMERORDER <= 2 " & _
            " ORDER BY CUSTOMERORDER"
  
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = objIADOAssist.GetConnStr
    cmd.CommandText = strSQL
    Set rst = cmd.Execute(1, , adCmdText)

    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadCustomerRoleData()

    Do While Not rst.EOF
        xmlDoc.firstChild.appendChild objIDOAssist.GetXMLFromRecordSet(rst, xmlClassDefDoc)
        rst.MoveNext
    Loop

        rst.Close

    Set IApplicationDO_FindMainApplicants = xmlDoc.documentElement

    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objErrAssist = Nothing
    Set objIDOAssist = Nothing
    Set objIADOAssist = Nothing
    Set objXmlAssist = Nothing
    Set objSqlAssist = Nothing
    Set cmd = Nothing
    Set rst = Nothing
    Set xmlDoc = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

FindMainApplicantsVbErr:
    
    ' Close the recordset if still open
        If Not rst Is Nothing Then
                If rst.State = adStateOpen Then
                        rst.Close
                End If
        End If

    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objIADOAssist = Nothing
    Set cmd = Nothing
    Set rst = Nothing
    Set xmlDoc = Nothing
    Set objXmlAssist = Nothing
    Set objSqlAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
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

Private Function IApplicationDO_FindMenuItemStatusList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode

On Error GoTo FindMenuItemStatusListVbErr

    Const strFunctionName As String = "IApplicationDO_FindMenuItemStatusList"

    Dim objErrAssist As New ErrAssist
    
    Dim objIDOAssist As IDOAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
         
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
     
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadMenuItemStatusData()
    Set IApplicationDO_FindMenuItemStatusList = objIDOAssist.FindListMultiple(vxmlTableElement, xmlClassDefDoc)

    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objErrAssist = Nothing
    Set objIDOAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

FindMenuItemStatusListVbErr:
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
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

Private Function IApplicationDO_FindOtherResidentList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for all instances of the persistant data associated with
'   this data object for the values supplied
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo FindOtherResidentListVbErr

    Const strFunctionName As String = "IApplicationDO_FindOtherResidentList"

    Dim objErrAssist As New ErrAssist
    
    Dim objIDOAssist As IDOAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
         
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
     
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadOtherResidentData()
    Set IApplicationDO_FindOtherResidentList = objIDOAssist.FindListMultiple(vxmlTableElement, xmlClassDefDoc)

    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objErrAssist = Nothing
    Set objIDOAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

FindOtherResidentListVbErr:
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
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

Private Function IApplicationDO_FindVerificationList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for all instances of the persistant data associated with
'   this data object for the values supplied
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo FindVerificationListVbErr

    Const strFunctionName As String = "IApplicationDO_FindVerificationList"

    Dim objErrAssist As New ErrAssist
    
    Dim objIDOAssist As IDOAssist, _
        objXmlAssist As New XMLAssist
    Dim objComboDO As New ComboDO
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
    
    Dim xmlVerificationList     As IXMLDOMNodeList, _
        xmlVerificationNode     As IXMLDOMNode, _
        xmlVerificationTypeNode As IXMLDOMElement

    Dim strVerificationType   As String, _
        strIdentificationType As String
    Dim strTemp As String

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
         
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
     
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadVerificationData
    Set IApplicationDO_FindVerificationList = objIDOAssist.FindListMultiple(vxmlTableElement, xmlClassDefDoc)

    '
    ' Get identification type string for each record
    '
    Set xmlVerificationList = vxmlTableElement.selectNodes("VERIFICATION")

    For Each xmlVerificationNode In xmlVerificationList
        ' Insert combo string for IdentificationType tag
        Set xmlVerificationTypeNode = xmlVerificationNode.selectSingleNode("VERIFICATIONTYPE")

        strVerificationType = objXmlAssist.GetTagValue(xmlVerificationNode, "VERIFICATIONTYPE")
        strIdentificationType = objXmlAssist.GetTagValue(xmlVerificationNode, "IDENTIFICATIONTYPE")
    
        If IsNumeric(strVerificationType) And IsNumeric(strIdentificationType) Then
            If CInt(strVerificationType) = 1 Then
                ' Insert from PersonalIdType combo
                strTemp = objComboDO.GetComboText("PersonalIDType", strIdentificationType)
            ElseIf CInt(strVerificationType) = 2 Then
                ' Insert from ResidencyIdType combo
                strTemp = objComboDO.GetComboText("ResidencyIDType", strIdentificationType)
            End If

            If strTemp <> "" Then
                xmlVerificationTypeNode.setAttribute "TEXT", strTemp
            End If
        End If
    Next xmlVerificationNode

    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objErrAssist = Nothing
    Set objIDOAssist = Nothing
    Set xmlVerificationList = Nothing
    Set xmlVerificationNode = Nothing
    Set xmlVerificationTypeNode = Nothing
    Set objComboDO = Nothing
    Set objXmlAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

FindVerificationListVbErr:
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set xmlVerificationList = Nothing
    Set xmlVerificationNode = Nothing
    Set xmlVerificationTypeNode = Nothing
    Set objComboDO = Nothing
    Set objXmlAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
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

Private Function IApplicationDO_FindUserHistoryList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for all instances of the persistant data associated with
'   this data object for the values supplied
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo FindUserHistoryListVbErr

    Const strFunctionName As String = "IApplicationDO_FindUserHistoryList"

    Dim objErrAssist As New ErrAssist
    
    Dim objIDOAssist As IDOAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
         
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
     
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoaduserHistoryData()
    Set IApplicationDO_FindUserHistoryList = objIDOAssist.FindListMultiple(vxmlTableElement, xmlClassDefDoc, "USERHISTORYDATE DESC")

    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objErrAssist = Nothing
    Set objIDOAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

FindUserHistoryListVbErr:
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
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

Private Function IApplicationDO_GetApplicationCreditCard(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   this data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo GetApplicationCreditCardVbErr

    Const strFunctionName As String = "IApplicationDO_GetApplicationCreditCard"
    
    Dim objErrAssist As New ErrAssist
    
    Dim objIDOAssist As IDOAssist, _
        objXmlAssist As New XMLAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
    Dim xmlReturn As IXMLDOMNode
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
            
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadApplicationCreditCardData
    Set xmlReturn = objIDOAssist.GetData(vxmlTableElement, xmlClassDefDoc, vstritemname:="APPLICATIONCREDITCARD")

    Set IApplicationDO_GetApplicationCreditCard = xmlReturn

    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlReturn = Nothing
    Set objXmlAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

GetApplicationCreditCardVbErr:
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set xmlReturn = Nothing
    Set objXmlAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
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

Private Function IApplicationDO_GetApplicationData(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the APPLICATION'S LATEST DETAILS. Note this fetches data from view 'ApplicationLatestDetails'.
'   This methos is different from GetDat
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo GetApplicationDataVbErr

    Const strFunctionName As String = "IApplicationDO_GetApplicationData"
    
    Dim objErrAssist As New ErrAssist
    
    Dim objIDOAssist As IDOAssist, _
        objIADOAssist As IADOAssist, _
        objSqlAssist As New SQLAssist
    Dim objXmlAssist As New XMLAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
    Dim xmlDoc As New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    
    Dim cmd As ADODB.Command, _
        rst As ADODB.Recordset
    Dim strApplicationNumber As String, _
        strSQL               As String
    ' SR 19/12/2002 : BM0209 - Start
    Dim blnCMDataOnly As Boolean
    Dim objConvertAssit As New ConvertAssist
    ' SR 19/12/2002 : BM0209 - End
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
        Set objIADOAssist = New ADOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
        Set objIADOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    End If

    ' Get application number from XML
    strApplicationNumber = objXmlAssist.GetTagValue(vxmlTableElement, "APPLICATIONNUMBER")
    If strApplicationNumber = "" Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingPrimaryTag
    End If

    ' SR 19/12/2002 : BM0209 - Get the data required for cost modelling and do not create locks on the tables
    blnCMDataOnly = objConvertAssit.CSafeBool(objXmlAssist.GetTagValue(vxmlTableElement, "CALCSDATAONLY"))
    If blnCMDataOnly Then
        '*=[MC]BMIDS763 APPLICATION DATE COLUMN ADDED TO THE QUERY.
        'SR 13/08/2004:BMIDS815 - Add column LASTACCOUNTDOWNLOADDATE, also use ApplicationLatestDate from view rather than table Application
        'JD BMIDS975 15/03/2005 - Add return of FurtherAdvanceTermInd.
        'MAR18 Add return of LevelOfAdvice
        'EP2_780 Add return of SPECIALSCHEME
        strSQL = " SELECT A.APPLICATIONNUMBER, A.APPLICATIONFACTFINDNUMBER, A.ACTIVEQUOTENUMBER, A.ACCEPTEDQUOTENUMBER, A.TYPEOFAPPLICATION,  A.PURCHASEPRICEORESTIMATEDVALUE,A.CHANNELID, A.APPLICATIONDATE, " & _
                " A.LASTACCOUNTDOWNLOADDATE, A.FURTHERADVANCETERMIND, A.LEVELOFADVICE , A.SPECIALSCHEME" & _
                " FROM APPLICATIONLATESTDETAILS  A WITH (NOLOCK) " & _
                " WHERE  A.APPLICATIONNUMBER = " & objSqlAssist.FormatString(strApplicationNumber)
        'SECTION END
        'SR 13/08/2004:BMIDS815 - End
                
    Else
    ' SR 19/12/2002 : BM0209 - End
        strSQL = "SELECT * FROM APPLICATIONLATESTDETAILS WHERE " & _
                " APPLICATIONNUMBER = " & objSqlAssist.FormatString(strApplicationNumber)
    End If ' SR 19/12/2002 : BM0209
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = objIADOAssist.GetConnStr
    cmd.CommandText = strSQL

    Set rst = cmd.Execute(1, , adCmdText)

    Set objIClassDef = New ApplicationClassDef
    
    ' SR 19/12/2002 : BM0209 - Start
    If blnCMDataOnly Then
       Set xmlClassDefDoc = objIClassDef.LoadApplicationDataForCostModelling
    Else ' SR 19/12/2002 : BM0209 - End
        Set xmlClassDefDoc = objIClassDef.LoadApplicationLatestData
    End If ' SR 19/12/2002 : BM0209
    
    
    ' KRW BMIDS888 22/11/2004
    'xmlDoc.appendChild objIDOAssist.GetXMLFromRecordset(rst, xmlClassDefDoc) 'old
    xmlDoc.appendChild objIDOAssist.GetXMLFromRecordSet(rst, xmlClassDefDoc).cloneNode(True) 'new
    
    
        rst.Close
        
    ' SR 19/12/2002 : BM0209 - Start
    If Not blnCMDataOnly Then
        objXmlAssist.PromoteNode xmlDoc.documentElement, xmlDoc.documentElement.selectSingleNode("APPLICATION")
    Else ' SR 19/12/2002 : BM0209 - End
        objXmlAssist.ChangeNodeName xmlDoc.documentElement, "APPLICATIONLATESTDETAILS", "APPLICATIONFACTFIND"
    End If ' SR 19/12/2002 : BM0209

    ' KRW BMIDS888 22/11/2004 took out above processing
'    If blnCMDataOnly Then
'        objXmlAssist.ChangeNodeName xmlDoc.documentElement, "APPLICATIONLATESTDETAILS", "APPLICATIONFACTFIND"
'   End If






    Set IApplicationDO_GetApplicationData = xmlDoc.documentElement
 
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objSqlAssist = Nothing
    Set cmd = Nothing
    Set rst = Nothing
    
' JR SYS2048
'    If Not objContext Is Nothing Then
'        objContext.SetComplete
'    End If
    
    Exit Function

GetApplicationDataVbErr:
    
    ' Close the recordset if still open
        If Not rst Is Nothing Then
                If rst.State = adStateOpen Then
                        rst.Close
                End If
        End If

    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objXmlAssist = Nothing
    Set objSqlAssist = Nothing
    Set cmd = Nothing
    Set rst = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
'JR SYS2048
'    If Not objContext Is Nothing Then
'        If objErrAssist.IsSystemError Then
'            objContext.SetAbort
'        Else
'            objContext.SetComplete
'        End If
'    End If
    
    Set objErrAssist = Nothing
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Function IApplicationDO_GetApplicationFactFindData(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant from table 'ApplicationFactFind'
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------

On Error GoTo GetApplicationFactFindDataVbErr

    Const strFunctionName As String = "IApplicationDO_GetApplicationFactFindData"
    
    Dim objErrAssist As New ErrAssist
    
    Dim objIDOAssist As IDOAssist
    Dim objXmlAssist As New XMLAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
            
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadApplicationFactFindData

    Set IApplicationDO_GetApplicationFactFindData = objIDOAssist.GetData(vxmlTableElement, xmlClassDefDoc)

    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

GetApplicationFactFindDataVbErr:
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objXmlAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
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

Private Function IApplicationDO_GetApplicationOfferData(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   the table 'ApplicationOffer'
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo GetApplicationOfferDataVbErr

    Const strFunctionName As String = "IApplicationDO_GetApplicationOfferData"
    
    Dim objErrAssist As New ErrAssist
    Dim conn As ADODB.Connection
    Dim rst As ADODB.Recordset
    
    Dim strSQL As String
    Dim strApplicationNumber As String
    Dim strApplicationFFNumber As String
    
    Dim objIDOAssist As IDOAssist
    Dim objIADOAssist As IADOAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
    Dim objXmlAssist As New XMLAssist
    Dim objResponseNode As IXMLDOMNode
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
            
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
        Set objIADOAssist = New ADOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
        Set objIADOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    End If
    
    Set objResponseNode = vxmlTableElement.ownerDocument.createElement("RESPONSE")

    strApplicationNumber = objXmlAssist.GetMandatoryElementText(vxmlTableElement, ".//APPLICATIONNUMBER")
    strApplicationFFNumber = objXmlAssist.GetMandatoryElementText(vxmlTableElement, ".//APPLICATIONFACTFINDNUMBER")
        
    strSQL = "select * from APPLICATIONOFFER"
    strSQL = strSQL & " where applicationnumber='" & strApplicationNumber
    strSQL = strSQL & "' and applicationfactfindnumber ='" & strApplicationFFNumber
    strSQL = strSQL & "' and offersequenceno = (SELECT MAX(offersequenceno)FROM APPLICATIONOFFER" ' KRW 25/08/04 BMIDS810
    strSQL = strSQL & " WHERE applicationnumber='" & strApplicationNumber
    strSQL = strSQL & "' and applicationfactfindnumber = '" & strApplicationFFNumber & "')"
     
    'Get Database Connection
    Set conn = New ADODB.Connection
    conn.ConnectionString = objIADOAssist.GetConnStr
    conn.open
        
    'Open recordset
    Set rst = New ADODB.Recordset
    rst.CursorLocation = adUseClient
    rst.open strSQL, conn, adOpenStatic, adLockReadOnly, adCmdText
    
    If rst.BOF And rst.EOF Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound
    End If
    
    'Get XML from the recordset
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadApplicationOfferData
    objIDOAssist.GetXMLFromRecordSet rst, xmlClassDefDoc, objResponseNode
    rst.Close
            
    'Disconnect recordset
    Set rst.ActiveConnection = Nothing
        conn.Close
         
    Set IApplicationDO_GetApplicationOfferData = objResponseNode

    Set rst = Nothing
    Set conn = Nothing
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objIADOAssist = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objResponseNode = Nothing
    
    'JR SYS2048
'    If Not objContext Is Nothing Then
'        objContext.SetComplete
'    End If
    
    Exit Function

GetApplicationOfferDataVbErr:
    
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
        Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objIADOAssist = Nothing
    Set objXmlAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
' JR SYS2048
'    If Not objContext Is Nothing Then
'        If objErrAssist.IsSystemError Then
'            objContext.SetAbort
'        Else
'            objContext.SetComplete
'        End If
'    End If
    
    Set objErrAssist = Nothing
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function

Private Function IApplicationDO_GetApplicationPriority(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   the table 'ApplicationStage'
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo GetApplicationPriorityVbErr

    Const strFunctionName As String = "IApplicationDO_GetApplicationPriority"
    
    Dim objIDOAssist As IDOAssist
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim objIADOAssist As IADOAssist
    Dim objSqlAssist As New SQLAssist
    
    Dim xmlDoc As New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    Dim xmlElem As IXMLDOMElement
    Dim strConnString As String
    Dim strApplicationNumber As String
    Dim strSQL As String
    
    Dim adoConnection As ADODB.Connection
    Dim adoRecordset As ADODB.Recordset
    Dim adoCommand As ADODB.Command
    Dim adoParameter As ADODB.Parameter
       
    Set adoConnection = New ADODB.Connection
    Set adoRecordset = New ADODB.Recordset
    Set adoCommand = New ADODB.Command
    Set adoParameter = New ADODB.Parameter
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
            
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
        Set objIADOAssist = New ADOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
        Set objIADOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    End If
    
    strApplicationNumber = objXmlAssist.GetTagValue(vxmlTableElement, "APPLICATIONNUMBER")
    
    strSQL = ""
    strSQL = strSQL + " SELECT  * FROM APPLICATIONPRIORITY "
    strSQL = strSQL + " WHERE APPLICATIONNUMBER = ? "
    strSQL = strSQL + " AND APPPRIORITYSEQUENCENO = ( SELECT MAX(APPPRIORITYSEQUENCENO) FROM APPLICATIONPRIORITY "
    strSQL = strSQL + " WHERE APPLICATIONNUMBER = ? ) "

    adoCommand.CommandText = strSQL
    
    ' Set up the parameters to be substituted into the SQL
    Set adoParameter = adoCommand.CreateParameter("ApplicationNumber", adBSTR, adParamInput, Len(strApplicationNumber), strApplicationNumber)
    adoCommand.Parameters.Append adoParameter
    
    Set adoParameter = adoCommand.CreateParameter("MaxApplicationNumber", adBSTR, adParamInput, Len(strApplicationNumber), strApplicationNumber)
    adoCommand.Parameters.Append adoParameter
    
    ' Set Recordset properties. Must be adUseClient for disconnected recordset
    adoRecordset.CursorLocation = adUseClient
    adoRecordset.CursorType = adOpenForwardOnly
    adoRecordset.LockType = adLockReadOnly
    
    strConnString = objIADOAssist.GetConnStr()
    adoConnection.open strConnString
    
    adoCommand.ActiveConnection = adoConnection
    Set adoRecordset.Source = adoCommand

    'Populate recordset
    adoRecordset.open
    
    'Disconnect the recordset
    Set adoRecordset.ActiveConnection = Nothing
    adoConnection.Close
    
    If adoRecordset.EOF Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound
    End If
    
    ' APS SYS1993 Change to return all columns on the recordset
    'Set xmlElem = xmlDoc.createElement("APPLICATIONPRIORITY")
    'xmlElem.Text = adoRecordset.Fields("APPLICATIONPRIORITYVALUE").Value
    
    Dim objClassDef As IApplicationClassDef
    Set objClassDef = New ApplicationClassDef
    
    Dim xmlClassDoc As New FreeThreadedDOMDocument40
    xmlClassDoc.validateOnParse = False
    xmlClassDoc.setProperty "NewParser", True
    Set xmlClassDoc = objClassDef.LoadApplicationPriorityData()
    
    Dim xmlNode As IXMLDOMNode
    Set xmlNode = objIDOAssist.GetXMLFromRecordSet(adoRecordset, xmlClassDoc)
    
    'Populate recordset
    adoRecordset.Close
    
    Set IApplicationDO_GetApplicationPriority = xmlNode
    
    Set xmlNode = Nothing
    Set objIDOAssist = Nothing
    Set objIADOAssist = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objClassDef = Nothing
    Set xmlClassDoc = Nothing
    Set adoConnection = Nothing
    Set adoCommand = Nothing
    Set adoRecordset = Nothing
    Set adoParameter = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

GetApplicationPriorityVbErr:
    
    If adoConnection.State = adStateOpen Then
        adoConnection.Close
    End If
    
    If Not adoRecordset Is Nothing Then
        If adoRecordset.State = adStateOpen Then
            adoRecordset.Close
        End If
    End If
    
    Set xmlNode = Nothing
    Set objClassDef = Nothing
    Set xmlClassDoc = Nothing
    Set objIDOAssist = Nothing
    Set objIADOAssist = Nothing
    Set objXmlAssist = Nothing
    Set adoConnection = Nothing
    Set adoCommand = Nothing
    Set adoRecordset = Nothing
    Set adoParameter = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
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

Private Function IApplicationDO_IsAppUnderReview(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   the table 'ApplicationStage'
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo IsAppUnderReviewVbErr

    Const strFunctionName As String = "IApplicationDO_IsAppUnderReview"
    
    Dim objErrAssist As New ErrAssist
    Dim objIADOAssist As IADOAssist
    Dim objSqlAssist As New SQLAssist
    Dim objXmlAssist As New XMLAssist
    Dim objIDOAssist As IDOAssist
    
    Dim xmlDoc As New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    Dim xmlElem As IXMLDOMElement
    
    Dim adoConnection As ADODB.Connection
    Dim adoRecordset As ADODB.Recordset
    Dim adoCommand As ADODB.Command
    Dim adoParameter As ADODB.Parameter
    
    Set adoConnection = New ADODB.Connection
    Set adoRecordset = New ADODB.Recordset
    Set adoCommand = New ADODB.Command
    Set adoParameter = New ADODB.Parameter
    
    Dim strSQL As String
    Dim strConnString As String
    Dim strApplicationNumber As String
    Dim strApplicationFactFindNumber As String
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
        Set objIADOAssist = New ADOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
        Set objIADOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    End If
    
    strApplicationNumber = objXmlAssist.GetTagValue(vxmlTableElement, "APPLICATIONNUMBER")
    strApplicationFactFindNumber = objXmlAssist.GetTagValue(vxmlTableElement, "APPLICATIONFACTFINDNUMBER")
        
    strSQL = ""
    strSQL = strSQL + " SELECT COUNT(1) FROM APPLICATIONREVIEWHISTORY "
    strSQL = strSQL + " WHERE APPLICATIONNUMBER = ? "
    strSQL = strSQL + " AND APPLICATIONFACTFINDNUMBER = ? "
    strSQL = strSQL + " AND OVERRIDEDATETIME IS NULL "
    
    adoCommand.CommandText = strSQL
    
    ' Set up the parameters to be substituted into the SQL
    Set adoParameter = adoCommand.CreateParameter("ApplicationNumber", adBSTR, adParamInput, Len(strApplicationNumber), strApplicationNumber)
    adoCommand.Parameters.Append adoParameter
    
    Set adoParameter = adoCommand.CreateParameter("ApplicationFactFindNumber", adBSTR, adParamInput, Len(strApplicationFactFindNumber), strApplicationFactFindNumber)
    adoCommand.Parameters.Append adoParameter
    
    ' Set Recordset properties. Must be adUseClient for disconnected recordset
    adoRecordset.CursorLocation = adUseClient
    adoRecordset.CursorType = adOpenForwardOnly
    adoRecordset.LockType = adLockReadOnly
    
    strConnString = objIADOAssist.GetConnStr()
    adoConnection.open strConnString
    
    adoCommand.ActiveConnection = adoConnection
    Set adoRecordset.Source = adoCommand

    'Populate recordset
    adoRecordset.open
    
    'Disconnect the recordset
    Set adoRecordset.ActiveConnection = Nothing
    adoConnection.Close
    
    If adoRecordset.EOF Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound
    End If
    
    Set xmlElem = xmlDoc.createElement("UNDERREVIEWINDICATOR")
    
    If adoRecordset.fields(0).Value >= 1 Then
        xmlElem.Text = 1
    Else
        xmlElem.Text = 0
    End If
    
    Set IApplicationDO_IsAppUnderReview = xmlElem
    
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set adoConnection = Nothing
    Set adoCommand = Nothing
    Set adoRecordset = Nothing
    Set adoParameter = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

IsAppUnderReviewVbErr:
    
    Set objIDOAssist = Nothing
    Set objXmlAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
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

Private Function IApplicationDO_GetApplicationStage(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   the table 'ApplicationStage'
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo GetApplicationStageVbErr

    Const strFunctionName As String = "IApplicationDO_GetApplicationStage"
    
    Dim objErrAssist As New ErrAssist
    
    Dim objIDOAssist As IDOAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
    Dim objXmlAssist As New XMLAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
            
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadApplicationStageData
    Set IApplicationDO_GetApplicationStage = objIDOAssist.GetData(vxmlTableElement, xmlClassDefDoc)

    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

GetApplicationStageVbErr:
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objXmlAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
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

Private Function IApplicationDO_GetApplicationVerification(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   this data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo GetApplicationVerificationVbErr

    Const strFunctionName As String = "IApplicationDO_GetApplicationVerification"
    
    Dim objErrAssist As New ErrAssist
    
    Dim objIDOAssist As IDOAssist, _
        objXmlAssist As New XMLAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
    Dim xmlReturn As IXMLDOMNode
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
            
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadApplicationVerificationData
    Set xmlReturn = objIDOAssist.GetData(vxmlTableElement, xmlClassDefDoc, vstritemname:="APPLICATIONVERIFICATION")

    Set IApplicationDO_GetApplicationVerification = xmlReturn

    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlReturn = Nothing
    Set objXmlAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

GetApplicationVerificationVbErr:
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set xmlReturn = Nothing
    Set objXmlAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
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


Private Function IApplicationDO_GetCurrentApplicationStage(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a latest instance of the persistant data associated with
'   the table 'ApplicationStage'
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo GetCurrentApplicationStageVbErr

    Const strFunctionName As String = "IApplicationDO_GetCurrentApplicationStage"
    
    Dim objErrAssist As New ErrAssist
    Dim conn As ADODB.Connection
    Dim rst As ADODB.Recordset
    
    Dim strSQL As String
    Dim strApplicationNumber As String
    Dim strApplicationFFNumber As String
    
    Dim objIDOAssist As IDOAssist
    Dim objIADOAssist As IADOAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
    Dim objXmlAssist As New XMLAssist
    Dim objResponseNode As IXMLDOMNode
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
            
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
        Set objIADOAssist = New ADOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
        Set objIADOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    End If
    
    Set objResponseNode = vxmlTableElement.ownerDocument.createElement("RESPONSE")

    strApplicationNumber = objXmlAssist.GetMandatoryElementText(vxmlTableElement, ".//APPLICATIONNUMBER")
    strApplicationFFNumber = objXmlAssist.GetMandatoryElementText(vxmlTableElement, ".//APPLICATIONFACTFINDNUMBER")
    
    ' APS 28/02/01 SYS1990
    ' PSC 09/01/2003 BM0089
    strSQL = "select * from APPLICATIONSTAGE WITH (NOLOCK)"
    strSQL = strSQL & " where applicationnumber='" & strApplicationNumber
    strSQL = strSQL & "' and applicationfactfindnumber = " & strApplicationFFNumber
    strSQL = strSQL & " and CaseSequenceNo = (SELECT Max(CASESEQUENCENO) from APPLICATIONSTAGE"
    strSQL = strSQL & " where applicationnumber='" & strApplicationNumber
    strSQL = strSQL & "' and applicationfactfindnumber = " & strApplicationFFNumber & ")"


    'Get Database Connection
    Set conn = New ADODB.Connection
    conn.ConnectionString = objIADOAssist.GetConnStr
    conn.open
        
    'Open recordset
    Set rst = New ADODB.Recordset
    rst.CursorLocation = adUseClient
    rst.open strSQL, conn, adOpenStatic, adLockReadOnly, adCmdText
    
    
    
    If rst.BOF And rst.EOF Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound
    End If
    
    'Get XML from the recordset
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadApplicationStageData
    objIDOAssist.GetXMLFromRecordSet rst, xmlClassDefDoc, objResponseNode
    rst.Close
    
        
        'Disconnect recordset
    Set rst.ActiveConnection = Nothing
        conn.Close
         
    Set IApplicationDO_GetCurrentApplicationStage = objResponseNode

        Set rst = Nothing
        Set conn = Nothing
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objIADOAssist = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objResponseNode = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

GetCurrentApplicationStageVbErr:
    
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
        Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objIADOAssist = Nothing
    Set objXmlAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
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

Private Function IApplicationDO_GetCustomerRelationship(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   this data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo GetCustomerRelationshipVbErr

    Const strFunctionName As String = "IApplicationDO_GetCustomerRelationship"
    
    Dim objErrAssist As New ErrAssist
    
    Dim objIDOAssist As IDOAssist, _
        objXmlAssist As New XMLAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
    
    Dim xmlMainApplicants    As IXMLDOMNode, _
        xmlMainApplicantList As IXMLDOMNodeList, _
        xmlCRNode            As IXMLDOMNode
    Dim xmlNode As IXMLDOMNode

    Dim strCustomerNumber1        As String, _
        strCustomerVersionNumber1 As String, _
        strCustomerNumber2        As String, _
        strCustomerVersionNumber2 As String

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
            
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If

    If vxmlTableElement.nodeName = "CUSTOMERRELATIONSHIP" Then
        Set xmlCRNode = vxmlTableElement
    Else
        Set xmlCRNode = vxmlTableElement.selectSingleNode("CUSTOMERRELATIONSHIP")
    End If

    If xmlCRNode Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingPrimaryTag
    End If

    '
    ' Analyse main applicants
    '
    Set xmlMainApplicants = IApplicationDO_FindMainApplicants(vxmlTableElement)
    Set xmlMainApplicantList = xmlMainApplicants.selectNodes("CUSTOMERROLE")

    If xmlMainApplicantList.length > 1 Then
        strCustomerNumber1 = objXmlAssist.GetTagValue(xmlMainApplicantList.Item(0), "CUSTOMERNUMBER")
        strCustomerVersionNumber1 = objXmlAssist.GetTagValue(xmlMainApplicantList.Item(0), "CUSTOMERVERSIONNUMBER")
        strCustomerNumber2 = objXmlAssist.GetTagValue(xmlMainApplicantList.Item(1), "CUSTOMERNUMBER")
        strCustomerVersionNumber2 = objXmlAssist.GetTagValue(xmlMainApplicantList.Item(1), "CUSTOMERVERSIONNUMBER")
    Else
        ' Fewer than 2 main applicants
        objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound, "Expected TWO main applicants"
    End If

    '
    ' Update request XML to reflect the new search criteria
    '
    ' OwnerCustomerNumber
    Set xmlNode = xmlCRNode.selectSingleNode("OWNERCUSTOMERNUMBER")
    If xmlNode Is Nothing Then
        Set xmlNode = xmlCRNode.appendChild(xmlCRNode.ownerDocument.createElement("OWNERCUSTOMERNUMBER"))
    End If
    xmlNode.Text = strCustomerNumber1

    ' OwnerCustomerVersionNumber
    Set xmlNode = xmlCRNode.selectSingleNode("OWNERCUSTOMERVERSIONNUMBER")
    If xmlNode Is Nothing Then
        Set xmlNode = xmlCRNode.appendChild(xmlCRNode.ownerDocument.createElement("OWNERCUSTOMERVERSIONNUMBER"))
    End If
    xmlNode.Text = strCustomerVersionNumber1

    ' RelToCustomerNumber
    Set xmlNode = xmlCRNode.selectSingleNode("RELTOCUSTOMERNUMBER")
    If xmlNode Is Nothing Then
        Set xmlNode = xmlCRNode.appendChild(xmlCRNode.ownerDocument.createElement("RELTOCUSTOMERNUMBER"))
    End If
    xmlNode.Text = strCustomerNumber2

    ' RelToCustomerVersionNumber
    Set xmlNode = xmlCRNode.selectSingleNode("RELTOCUSTOMERVERSIONNUMBER")
    If xmlNode Is Nothing Then
        Set xmlNode = xmlCRNode.appendChild(xmlCRNode.ownerDocument.createElement("RELTOCUSTOMERVERSIONNUMBER"))
    End If
    xmlNode.Text = strCustomerVersionNumber2
    
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadCustomerRelationshipData
    Set IApplicationDO_GetCustomerRelationship = objIDOAssist.GetData(vxmlTableElement, xmlClassDefDoc)
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlMainApplicants = Nothing
    Set xmlMainApplicantList = Nothing
    Set xmlCRNode = Nothing
    Set xmlNode = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

GetCustomerRelationshipVbErr:
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlMainApplicants = Nothing
    Set xmlMainApplicantList = Nothing
    Set xmlCRNode = Nothing
    Set xmlNode = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
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

Private Function IApplicationDO_GetData(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   this data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------

On Error GoTo GetDataVbErr

    Const strFunctionName As String = "IApplicationDO_GetData"
    
    Dim objErrAssist As New ErrAssist
    
    Dim objIDOAssist As IDOAssist
    Dim objXmlAssist As New XMLAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
            
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadApplicationData

    Set IApplicationDO_GetData = objIDOAssist.GetData(vxmlTableElement, xmlClassDefDoc)

    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

GetDataVbErr:
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objXmlAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
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

Private Function IApplicationDO_GetDependantForCustomer(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   this data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo GetDependantForCustomerVbErr

    Const strFunctionName As String = "IApplicationDO_GetDependantForCustomer"
    
    Dim objErrAssist As New ErrAssist
    
    Dim objIDOAssist As IDOAssist
    Dim objXmlAssist As New XMLAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
    Dim xmlReturn As IXMLDOMNode
    Dim xmlORRequest  As IXMLDOMNode, _
        xmlORResponse As IXMLDOMNode
    Dim xmlOtherResident As IXMLDOMNode
    
    Dim strApplicationNumber         As String, _
        strApplicationFactFindNumber As String, _
        strPersonGUID                As String

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
            
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadDependantDetailsData
    Set xmlReturn = objIDOAssist.GetData(vxmlTableElement, xmlClassDefDoc, vstrTableName:="DEPENDANT")

    objXmlAssist.PromoteNode xmlReturn, xmlReturn.selectSingleNode(".//DEPENDANT")

    '
    ' Get other resident details (if they exist)
    '
    strApplicationNumber = objXmlAssist.GetTagValue(vxmlTableElement, "APPLICATIONNUMBER")
    strApplicationFactFindNumber = objXmlAssist.GetTagValue(vxmlTableElement, "APPLICATIONFACTFINDNUMBER")
    strPersonGUID = objXmlAssist.GetTagValue(xmlReturn, "PERSONGUID")

    If (strApplicationNumber = "") Or (strApplicationFactFindNumber = "") Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingPrimaryTag
    End If

    If strPersonGUID <> "" Then
        Set xmlORRequest = vxmlTableElement.ownerDocument.createElement("OTHERRESIDENT")
        xmlORRequest.appendChild(vxmlTableElement.ownerDocument.createElement("APPLICATIONNUMBER")).Text _
            = strApplicationNumber
        xmlORRequest.appendChild(vxmlTableElement.ownerDocument.createElement("APPLICATIONFACTFINDNUMBER")).Text _
            = strApplicationFactFindNumber
        xmlORRequest.appendChild(vxmlTableElement.ownerDocument.createElement("PERSONGUID")).Text _
            = strPersonGUID
        On Error GoTo RecordNotFoundErr
        Set xmlORResponse = IApplicationDO_FindOtherResidentList(xmlORRequest)
        On Error GoTo GetDependantForCustomerVbErr
    End If

    '
    ' Append the other resident data to the dependant data as appropriate
    '
    If Not (xmlORResponse Is Nothing) Then
        Set xmlOtherResident = xmlORResponse.selectSingleNode("OTHERRESIDENTLIST/OTHERRESIDENT")
        If Not (xmlOtherResident Is Nothing) Then
            xmlReturn.appendChild xmlOtherResident
        End If
    End If

    Set IApplicationDO_GetDependantForCustomer = xmlReturn

    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlORRequest = Nothing
    Set xmlORResponse = Nothing
    Set xmlOtherResident = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

RecordNotFoundErr:
    If objErrAssist.GetOmigaErrorNumber(Err.Number) = oeRecordNotFound Then
        Resume Next
    End If

GetDependantForCustomerVbErr:
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlORRequest = Nothing
    Set xmlORResponse = Nothing
    Set xmlOtherResident = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
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

Private Function IApplicationDO_GetFinancialSummary(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   this data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo GetFinancialSummaryVbErr

    Const strFunctionName As String = "IApplicationDO_GetFinancialSummary"
    
    Dim objErrAssist As New ErrAssist
    
    Dim objIDOAssist As IDOAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
            
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadFinancialSummaryData
    Set IApplicationDO_GetFinancialSummary = objIDOAssist.GetData(vxmlTableElement, xmlClassDefDoc)

    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

GetFinancialSummaryVbErr:
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
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

Private Function IApplicationDO_GetLoanPropertyDetails(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   this data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo GetLoanPropertyDetailsVbErr

    Const strFunctionName As String = "IApplicationDO_GetLoanPropertyDetails"
    
    Dim objErrAssist As New ErrAssist
    
    Dim objIDOAssist As IDOAssist, _
        objXmlAssist As New XMLAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
    Dim xmlReturn As IXMLDOMNode
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
            
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadLoanPropertyDetailsData
    Set xmlReturn = objIDOAssist.GetData(vxmlTableElement, xmlClassDefDoc, vstritemname:="LOANPROPERTYDETAILS")

    Set IApplicationDO_GetLoanPropertyDetails = xmlReturn

    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlReturn = Nothing
    Set objXmlAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

GetLoanPropertyDetailsVbErr:
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set xmlReturn = Nothing
    Set objXmlAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
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

Private Function IApplicationDO_GetMortgageCalcDetails(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   the table 'MortgageCalculatorDetails'
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo GetMortgageCalcDetailsVbErr

    Const strFunctionName As String = "IApplicationDO_GetMortgageCalcDetails"
    
    Dim objErrAssist As New ErrAssist
    
    Dim objIDOAssist As IDOAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
    Dim objXmlAssist As New XMLAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
            
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadMortgageCalcDetailsData
    Set IApplicationDO_GetMortgageCalcDetails = objIDOAssist.GetData(vxmlTableElement, xmlClassDefDoc)

    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

GetMortgageCalcDetailsVbErr:
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objXmlAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
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

Private Function IApplicationDO_GetNewLoan(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   this data object
' pass:
'   vxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo GetNewLoanVbErr

    Const strFunctionName As String = "IApplicationDO_GetNewLoan"
    
    Dim objErrAssist As New ErrAssist
    
    Dim objIDOAssist As IDOAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
    Dim objXmlAssist As New XMLAssist
        
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadNewLoanData()
    Set IApplicationDO_GetNewLoan = objIDOAssist.GetData(vxmlTableElement, xmlClassDefDoc)

   
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

GetNewLoanVbErr:
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objXmlAssist = Nothing
    
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
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

Private Function IApplicationDO_GetOtherInsuranceCompany(ByVal vobjRequestNode As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   this data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo GetOtherInsuranceCompanyVbErr

    Const strFunctionName As String = "IApplicationDO_GetOtherInsuranceCompany"
    
    Dim objErrAssist As New ErrAssist
    
    Dim objIDOAssist As IDOAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
    Dim objXmlAssist As New XMLAssist
    Dim objADOAssist As ADOAssist
    Dim objSqlAssist As New SQLAssist
    
    Dim xmlReturn As IXMLDOMNode
    Dim xmlElement As IXMLDOMElement
    Dim xmlDocument As FreeThreadedDOMDocument40
    
    Dim strSQL As String
    Dim objConn As ADODB.Connection
    Dim objRs As ADODB.Recordset
    Dim strApplicationNumber As String
    Dim strApplicationFactFindNumber As String
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
            
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadOtherInsuranceCompanyData
    Set objADOAssist = New ADOAssist
    
    'Set xmlReturn = objIDOAssist.GetData(vxmlTableElement, xmlClassDefDoc, vstrTableName:="OTHERINSURANCECOMPANY")
    
    strApplicationNumber = objXmlAssist.GetMandatoryElementText(vobjRequestNode, "SEARCH/APPLICATIONNUMBER")
    strApplicationNumber = objSqlAssist.FormatString(strApplicationNumber)
    strApplicationFactFindNumber = objXmlAssist.GetMandatoryElementText(vobjRequestNode, "SEARCH/APPLICATIONFACTFINDNUMBER")
    
    'Get connection and find records required
    strSQL = "SELECT * FROM OTHERINSURANCECOMPANY WHERE APPLICATIONNUMBER = " & strApplicationNumber & " AND APPLICATIONFACTFINDNUMBER = " & strApplicationFactFindNumber
    Set objConn = New ADODB.Connection
    Set objRs = New ADODB.Recordset
    objConn.ConnectionString = objADOAssist.GetConnStr
    objConn.open
    objRs.CursorLocation = adUseClient
    objRs.open strSQL, objConn, adOpenStatic, adLockReadOnly, adCmdText
    
    If objRs.EOF Then
        'No record currently exists so need to create default response
        Set xmlDocument = New FreeThreadedDOMDocument40
        xmlDocument.validateOnParse = False
        xmlDocument.setProperty "NewParser", True
        Set xmlElement = xmlDocument.createElement("NEWINSURANCECOMPANY")
        Set xmlReturn = xmlDocument.appendChild(xmlElement)
        xmlReturn.Text = "1"
    Else
        'Create XML from the returned recordset
        Set xmlReturn = objIDOAssist.GetXMLFromRecordSet(objRs, xmlClassDefDoc, xmlReturn)
        objXmlAssist.PromoteNode xmlReturn, xmlReturn.selectSingleNode("OTHERINSURANCECOMPANY")
    End If
        objRs.Close

        'Disconnect recordset
    Set objRs.ActiveConnection = Nothing
        objConn.Close

    Set IApplicationDO_GetOtherInsuranceCompany = xmlReturn

        Set objRs = Nothing
        Set objConn = Nothing
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlReturn = Nothing
    Set objXmlAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

GetOtherInsuranceCompanyVbErr:
    
    ' Close the recordset if still open
        If Not objRs Is Nothing Then
                If objRs.State = adStateOpen Then
                        objRs.Close
                End If
        End If

    ' Close the connection if still open
    If Not objConn Is Nothing Then
                If objConn.State = adStateOpen Then
                        objConn.Close
                End If
        End If
            
        Set objRs = Nothing
    Set objConn = Nothing
        Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set xmlReturn = Nothing
    Set objXmlAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
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

Private Function IApplicationDO_GetOtherResident(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   this data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo GetOtherResidentVbErr

    Const strFunctionName As String = "IApplicationDO_GetOtherResident"
    
    Dim objErrAssist As New ErrAssist
    
    Dim objIDOAssist As IDOAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
    Dim objXmlAssist As New XMLAssist
    Dim xmlReturn As IXMLDOMNode
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
            
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadOtherResidentDetailsData
    Set xmlReturn = objIDOAssist.GetData(vxmlTableElement, xmlClassDefDoc, vstrTableName:="OTHERRESIDENT")

    objXmlAssist.PromoteNode xmlReturn, xmlReturn.selectSingleNode("OTHERRESIDENT")

    Set IApplicationDO_GetOtherResident = xmlReturn

    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set xmlReturn = Nothing
    Set objXmlAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

GetOtherResidentVbErr:
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set xmlReturn = Nothing
    Set objXmlAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
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

Private Function IApplicationDO_GetUnitFromApplication(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
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
On Error GoTo GetUnitFromApplicationVbErr

    Const strFunctionName As String = "IApplicationDO_GetUnitFromApplication"
    
            
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim objIClassDef As IClassDef
    Dim objIADOAssist As IADOAssist
    Dim xmlListElem As IXMLDOMElement
    
    Dim adoConnection As ADODB.Connection
    Dim adoCommand As ADODB.Command
    Dim adoRecordset As ADODB.Recordset
    Dim adoParameter As ADODB.Parameter
    Dim strApplicationNumber As String
    Dim strApplicationFFNumber As String
    Dim intRetries As Integer
    Dim intMaxAttempts As Integer
    Dim intAttempt As Integer
    Dim blnOpenedOk As Boolean
    Dim strConnString As String
    
        
    ' Check the correct keys have been passed in
    strApplicationNumber = m_objXmlAssist.GetNodeValue(vxmlTableElement, _
                                                     ".//APPLICATIONNUMBER")
        
    Set objIADOAssist = New ADOAssist
    
    Set adoConnection = New ADODB.Connection
    Set adoCommand = New ADODB.Command
    Set adoRecordset = New ADODB.Recordset
        
    Dim strSQL As String
          
    strSQL = "SELECT USERHISTORY.* FROM USERHISTORY, (SELECT APPLICATIONNUMBER, MIN(USERHISTORYDATE) MINDATE" _
             & " FROM USERHISTORY GROUP BY APPLICATIONNUMBER) USERHISTORYDATE WHERE USERHISTORY.APPLICATIONNUMBER=?" _
             & " AND USERHISTORYDATE.APPLICATIONNUMBER = USERHISTORY.APPLICATIONNUMBER AND" _
             & " USERHISTORYDATE.MINDATE = USERHISTORY.USERHISTORYDATE"
         
    adoCommand.CommandText = strSQL
    
    ' Set up the parameters to be substituted into the SQL
    Set adoParameter = adoCommand.CreateParameter("ApplicationNumber", adBSTR, adParamInput, Len(strApplicationNumber), strApplicationNumber)
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
    
        adoConnection.Errors.Clear
        
        adoConnection.open strConnString
        
        If adoConnection.Errors.Count = 0 Then
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
    adoRecordset.open
    
    'Disconnect the recordset
    Set adoRecordset.ActiveConnection = Nothing
    adoConnection.Close
    
    If adoRecordset.EOF Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound
    End If
     
    Set objIClassDef = New ClassDef
    Set xmlClassDefDoc = objIClassDef.LoadUnitData()
    
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True

    Set xmlListElem = xmlOut.createElement("UNITSUMMARY")
    xmlOut.appendChild xmlListElem

    ' Convert recordset to XML
    m_objIDOAssist.GetXMLFromWholeRecordset adoRecordset, xmlClassDefDoc, xmlListElem
    
    adoRecordset.Close
    
    Set IApplicationDO_GetUnitFromApplication = xmlListElem
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set xmlOut = Nothing
    Set objIADOAssist = Nothing
    Set xmlListElem = Nothing
    Set adoConnection = Nothing
    Set adoCommand = Nothing
    Set adoRecordset = Nothing
    Set adoParameter = Nothing
    
    m_objContext.SetComplete
    
    
    Exit Function

GetUnitFromApplicationVbErr:
    
    
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

Private Function IApplicationDO_GetUserHistory(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   this data object - Table 'UserHistory'
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo GetuserHistoryVbErr

    Const strFunctionName As String = "IApplicationDO_GetUserHistory"
    
    Dim objErrAssist As New ErrAssist
    
    Dim objIDOAssist As IDOAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
    Dim objXmlAssist As New XMLAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
            
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoaduserHistoryData
    Set IApplicationDO_GetUserHistory = objIDOAssist.GetData(vxmlTableElement, xmlClassDefDoc)

    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

GetuserHistoryVbErr:
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objXmlAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
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

Private Sub IApplicationDO_Update(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   UpdateApplication a single instance of the persistant data associated with this
'   data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo UpdateVbErr

    Const strFunctionName As String = "IApplicationDO_Update"
        
    Dim objErrAssist As New ErrAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
    Dim objIDOAssist As IDOAssist

    Dim objXmlAssist As New XMLAssist

    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If

    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadApplicationData()
    objIDOAssist.Update vxmlTableElement, xmlClassDefDoc
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Sub

UpdateVbErr:
        
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objXmlAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
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
End Sub

Private Sub IApplicationDO_UpdateApplication(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   UpdateApplication a single instance of the persistant data associated with this
'   data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo UpdateApplicationVbErr

    Const strFunctionName As String = "IApplicationDO_UpdateApplication"
        
    Dim objErrAssist As New ErrAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
    Dim objIDOAssist As IDOAssist

    Dim objXmlAssist As New XMLAssist
    Dim xmlAFFNode As IXMLDOMNode

    Dim strApplicationNumber As String

    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If

    '
    ' Insert APPLICATIONNUMBER into the APPLICATIONFACTFIND if necessary
    '
    strApplicationNumber = objXmlAssist.GetTagValue(vxmlTableElement, "APPLICATIONNUMBER")

    ' Insert an application number into the application fact find XML if necessary
    Set xmlAFFNode = vxmlTableElement.selectSingleNode("APPLICATIONFACTFIND")
    If Not (xmlAFFNode Is Nothing) Then
        ' Application fact find data has been specified
        If objXmlAssist.GetTagValue(xmlAFFNode, "APPLICATIONNUMBER") = "" Then
            ' Create the application number field for APPLICATIONFACTFIND
            xmlAFFNode.appendChild( _
                xmlAFFNode.ownerDocument.createElement("APPLICATIONNUMBER")).Text = strApplicationNumber
        End If
    End If
    
    '
    ' Update APPLICATION
    '
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadApplicationData()
    objIDOAssist.Update vxmlTableElement, xmlClassDefDoc

    '
    ' Update APPLICATIONFACTFIND
    '
    If Not (xmlAFFNode Is Nothing) Then
        Set xmlClassDefDoc = objIClassDef.LoadApplicationFactFindData()
        objIDOAssist.Update vxmlTableElement, xmlClassDefDoc
    End If
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlAFFNode = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Sub

UpdateApplicationVbErr:
        
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlAFFNode = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
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

End Sub

Private Sub IApplicationDO_UpdateApplicationCreditCard(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   Updates a single instance of the persistant data associated with the table
'  'ApplicationCreditCard'
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo UpdateApplicationCreditCardVbErr

    Const strFunctionName As String = "IApplicationDO_UpdateApplicationCreditCard"
        
    Dim objErrAssist As New ErrAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
    Dim objIDOAssist As IDOAssist

    If objContext Is Nothing Then
      Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If

    Set objIClassDef = New ApplicationClassDef
   
    Set xmlClassDefDoc = objIClassDef.LoadApplicationCreditCardData()
    objIDOAssist.Update vxmlTableElement, xmlClassDefDoc

    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Sub

UpdateApplicationCreditCardVbErr:
        
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
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


End Sub

Private Sub IApplicationDO_UpdateApplicationFactFind(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   Updates a single instance of the persistant data associated with the table
' 'ApplicatinFactFind'
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo UpdateApplicationFactFindVbErr

    Const strFunctionName As String = "IApplicationDO_UpdateApplicationFactFind"
        
    Dim objErrAssist As New ErrAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
    Dim objIDOAssist As IDOAssist

    If objContext Is Nothing Then
      Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadApplicationFactFindData()
    objIDOAssist.Update vxmlTableElement, xmlClassDefDoc
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Sub

UpdateApplicationFactFindVbErr:
        
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
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
End Sub

Private Sub IApplicationDO_UpdateApplicationStage(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   Updates a single instance of the persistant data associated with the table 'ApplicatinStage'
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo UpdateApplicationStageVbErr

    Const strFunctionName As String = "IApplicationDO_UpdateApplicationStage"
        
    Dim objErrAssist As New ErrAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
    Dim objIDOAssist As IDOAssist

    If objContext Is Nothing Then
      Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Call AddCurrentDateTimeToRequest("DATETIME", vxmlTableElement)
    
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadApplicationStageData()
    objIDOAssist.Update vxmlTableElement, xmlClassDefDoc
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Sub

UpdateApplicationStageVbErr:
        
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
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
End Sub

Private Sub IApplicationDO_UpdateApplicationUnderwriting(ByVal vxmlTableElement As IXMLDOMElement)
    
On Error GoTo UpdateApplicationUnderwritingVbErr

    Const strFunctionName As String = "IApplicationDO_UpdateApplicationUnderwriting"
        
    Dim objErrAssist As New ErrAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
    Dim objIDOAssist As IDOAssist

    If objContext Is Nothing Then
      Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If

    Set objIClassDef = New ApplicationClassDef
   
    Set xmlClassDefDoc = objIClassDef.LoadApplicationUnderwritingData()
    objIDOAssist.Update vxmlTableElement, xmlClassDefDoc

    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Sub

UpdateApplicationUnderwritingVbErr:
        
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
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

End Sub

Private Sub IApplicationDO_UpdateApplicationVerification(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   Updates a single instance of the persistant data associated with the table
'  'ApplicationVerification'
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo UpdateApplicationVerificationVbErr

    Const strFunctionName As String = "IApplicationDO_UpdateApplicationVerification"
        
    Dim objErrAssist As New ErrAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
    Dim objIDOAssist As IDOAssist

    If objContext Is Nothing Then
      Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If

    Set objIClassDef = New ApplicationClassDef
   
    Set xmlClassDefDoc = objIClassDef.LoadApplicationVerificationData()
    objIDOAssist.Update vxmlTableElement, xmlClassDefDoc

    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Sub

UpdateApplicationVerificationVbErr:
        
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
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

End Sub

Private Sub IApplicationDO_UpdateCustomerRelationship(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   UpdateCustomerRelationship a single instance of the persistant data associated with this
'   data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo UpdateCustomerRelationshipVbErr

    Const strFunctionName As String = "IApplicationDO_UpdateCustomerRelationship"
        
    Dim objErrAssist As New ErrAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
    Dim objIDOAssist As IDOAssist

    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadCustomerRelationshipData
    objIDOAssist.Update vxmlTableElement, xmlClassDefDoc
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Sub

UpdateCustomerRelationshipVbErr:
        
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
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

End Sub

Private Sub IApplicationDO_UpdateDeclaration(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   UpdateDeclaration a single instance of the persistant data associated with this
'   data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo UpdateDeclarationVbErr

    Const strFunctionName As String = "IApplicationDO_UpdateDeclaration"
        
    Dim objErrAssist As New ErrAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
    Dim objIDOAssist As IDOAssist

    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadDeclarationData
    objIDOAssist.Update vxmlTableElement, xmlClassDefDoc
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Sub

UpdateDeclarationVbErr:
        
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
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

End Sub

Private Sub IApplicationDO_UpdateFinancialSummary(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   UpdateFinancialSummary a single instance of the persistant data associated with this
'   data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo UpdateFinancialSummaryVbErr

    Const strFunctionName As String = "IApplicationDO_UpdateFinancialSummary"
        
    Dim objErrAssist As New ErrAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
    Dim objIDOAssist As IDOAssist

    If objContext Is Nothing Then
      Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadFinancialSummaryData
    objIDOAssist.Update vxmlTableElement, xmlClassDefDoc
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Sub

UpdateFinancialSummaryVbErr:
        
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
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

End Sub

Private Sub IApplicationDO_UpdateLoanProperty(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   Updates a single instance of the persistant data associated with the table
'  'MortgageCalculatorDetails'
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo UpdateLoanPropertyVbErr

    Const strFunctionName As String = "IApplicationDO_UpdateLoanProperty"
        
    Dim objErrAssist As New ErrAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
    Dim objIDOAssist As IDOAssist
    Dim objIADOAssist As IADOAssist 'SYS0662 MDC 02/05/2000
    Dim objXmlAssist As XMLAssist  'SYS0662 MDC 02/05/2000
    Dim xmlLeaseholdNode As IXMLDOMNode  ' JD MAR40

    Dim strApplicationNumber As String
    Dim strApplicationFactFindNumber As String
    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
        Set objIADOAssist = New ADOAssist   'SYS0662 MDC 02/05/2000
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
        Set objIADOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")    'SYS0662 MDC 02/05/2000
    End If

    Set objIClassDef = New ApplicationClassDef
    Set objXmlAssist = New XMLAssist   'SYS0662 MDC 02/05/2000
    
    ' NEWPROPERTY
    Set xmlClassDefDoc = objIClassDef.LoadNewPropertyData()
    objIDOAssist.Update vxmlTableElement, xmlClassDefDoc
    
    'Get AppNumber and AppFFNumber from XML
    strApplicationNumber = objXmlAssist.GetMandatoryElementText(vxmlTableElement, "NEWLOAN/APPLICATIONNUMBER")
    strApplicationFactFindNumber = objXmlAssist.GetMandatoryElementText(vxmlTableElement, "NEWLOAN/APPLICATIONFACTFINDNUMBER")
    
    ' NEWPROPERTYLEASEHOLD  JD MAR40 include leasehold data
    Set xmlClassDefDoc = objIClassDef.LoadNewPropertyLeaseholdData()
    'EP241 IK NEWPROPERTYLEASEHOLD is not mandatory
    Set xmlLeaseholdNode = m_objXmlAssist.GetNode(vxmlTableElement, "//NEWPROPERTY/NEWPROPERTYLEASEHOLD", False)
    If Not xmlLeaseholdNode Is Nothing Then
        If objIADOAssist.CheckRecordExists("NEWPROPERTYLEASEHOLD", "ApplicationNumber = '" & strApplicationNumber & "' AND ApplicationFactFindNumber = " & strApplicationFactFindNumber) = False Then
            objIDOAssist.Create xmlLeaseholdNode, xmlClassDefDoc
        Else
            objIDOAssist.Update xmlLeaseholdNode, xmlClassDefDoc
        End If
    End If


    ' NEWLOAN
    Set xmlClassDefDoc = objIClassDef.LoadNewLoanData()
    'SYS0662 MDC 02/05/2000
    If objIADOAssist.CheckRecordExists("NEWLOAN", "ApplicationNumber = '" & strApplicationNumber & "' AND ApplicationFactFindNumber = " & strApplicationFactFindNumber) = False Then
        objIDOAssist.Create vxmlTableElement, xmlClassDefDoc
    Else
        objIDOAssist.Update vxmlTableElement, xmlClassDefDoc
    End If

    ' APPLICATIONFACTFIND
    Set xmlClassDefDoc = objIClassDef.LoadApplicationFactFindData()
    objIDOAssist.Update vxmlTableElement, xmlClassDefDoc
    
    ' SHAREDOWNERSHIPDETAILS
    Set xmlClassDefDoc = objIClassDef.LoadSharedOwnershipDetailsData()
    'SYS0662 MDC 02/05/2000
    If objIADOAssist.CheckRecordExists("SHAREDOWNERSHIPDETAILS", "ApplicationNumber = '" & strApplicationNumber & "' AND ApplicationFactFindNumber = " & strApplicationFactFindNumber) = False Then
        objIDOAssist.Create vxmlTableElement, xmlClassDefDoc
    Else
        objIDOAssist.Update vxmlTableElement, xmlClassDefDoc
    End If

    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Sub

UpdateLoanPropertyVbErr:
        
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
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
End Sub

Private Sub IApplicationDO_UpdateMenuItemStatus(ByVal vxmlTableElement As IXMLDOMElement)

On Error GoTo UpdateMenuItemStatusVbErr

    Const strFunctionName As String = "IApplicationDO_UpdateMenuItemStatus"
        
    Dim objErrAssist As New ErrAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
    Dim objIDOAssist As IDOAssist

    If objContext Is Nothing Then
      Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadMenuItemStatusData()
    objIDOAssist.Update vxmlTableElement, xmlClassDefDoc
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Sub

UpdateMenuItemStatusVbErr:
        
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
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
End Sub

Private Sub IApplicationDO_UpdateMortgageCalcDetails(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   Updates a single instance of the persistant data associated with the table
'  'MortgageCalculatorDetails'
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo UpdateMortgageCalcDetailsVbErr

    Const strFunctionName As String = "IApplicationDO_UpdateMortgageCalcDetails"
        
    Dim objErrAssist As New ErrAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
    Dim objIDOAssist As IDOAssist

    If objContext Is Nothing Then
      Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadMortgageCalcDetailsData()
    objIDOAssist.Update vxmlTableElement, xmlClassDefDoc
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Sub

UpdateMortgageCalcDetailsVbErr:
        
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
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
End Sub

Private Sub IApplicationDO_UpdateNewLoan(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   Update a single instance of the persistant data associated with the table 'NewLoan'
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo UpdateNewLoanVbErr

    Const strFunctionName As String = "IApplicationDO_UpdateNewLoan"
            
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
    Dim objIDOAssist As IDOAssist

    If objContext Is Nothing Then
      Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadNewLoanData()
    objIDOAssist.Update vxmlTableElement, xmlClassDefDoc
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Sub

UpdateNewLoanVbErr:
        
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    Dim objErrAssist As New ErrAssist
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
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

End Sub

Private Sub IApplicationDO_UpdateOtherInsuranceCompany(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   Update a single instance of the persistant data associated with this
'   data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo UpdateOtherInsuranceCompanyVbErr

    Const strFunctionName As String = "IApplicationDO_UpdateOtherInsuranceCompany"
        
    Dim objErrAssist As New ErrAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
    Dim objIDOAssist As IDOAssist

    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
        'Set objPersonDO = New PersonDO
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
        'Set objPersonDO = objContext.CreateInstance(App.Title & ".PersonDO")
    End If

    ' Update OTHERINSURANCECOMPANY
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadOtherInsuranceCompanyData
    objIDOAssist.Update vxmlTableElement, xmlClassDefDoc
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Sub

UpdateOtherInsuranceCompanyVbErr:
        
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
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


End Sub

Private Sub IApplicationDO_UpdateOtherResident(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   UpdateOtherResident a single instance of the persistant data associated with this
'   data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo UpdateOtherResidentVbErr

    Const strFunctionName As String = "IApplicationDO_UpdateOtherResident"
        
    Dim objErrAssist As New ErrAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
    Dim objIDOAssist As IDOAssist
    Dim objPersonDO As IPersonDO

    Dim xmlPersonNode As IXMLDOMNode

    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
        Set objPersonDO = New PersonDO
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
        Set objPersonDO = objContext.CreateInstance(App.Title & ".PersonDO")
    End If

    '
    ' Update PERSON
    '
    Set xmlPersonNode = vxmlTableElement.selectSingleNode("PERSON")
    If Not (xmlPersonNode Is Nothing) Then
        objPersonDO.Update xmlPersonNode
    End If
    
    '
    ' Update OTHERRESIDENT
    '
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadOtherResidentData
    objIDOAssist.Update vxmlTableElement, xmlClassDefDoc
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    Set objPersonDO = Nothing
    Set xmlPersonNode = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Sub

UpdateOtherResidentVbErr:
        
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objPersonDO = Nothing
    Set xmlPersonNode = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
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

End Sub

Private Sub IApplicationDO_UpdateSharedOwnershipDetails(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   UpdateSharedOwnershipDetails a single instance of the persistant data associated with this
'   data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo UpdateSharedOwnershipDetailsVbErr

    Const strFunctionName As String = "IApplicationDO_UpdateSharedOwnershipDetails"
        
    Dim objErrAssist As New ErrAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
    Dim objIDOAssist As IDOAssist

    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadSharedOwnershipDetailsData
    objIDOAssist.Update vxmlTableElement, xmlClassDefDoc
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Sub

UpdateSharedOwnershipDetailsVbErr:
        
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
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
End Sub

Private Sub IApplicationDO_UpdateUserHistory(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   UpdateUserHistory a single instance of the persistant data associated with this
'   data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo UpdateUserHistoryVbErr

    Const strFunctionName As String = "IApplicationDO_UpdateUserHistory"
        
    Dim objErrAssist As New ErrAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
    Dim objIDOAssist As IDOAssist

    If objContext Is Nothing Then
      Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoaduserHistoryData()
    objIDOAssist.Update vxmlTableElement, xmlClassDefDoc
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Sub

UpdateUserHistoryVbErr:
        
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
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

End Sub

Private Sub IApplicationDO_UpdateVerification(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   UpdateVerification a single instance of the persistant data associated with this
'   data object
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo UpdateVerificationVbErr

    Const strFunctionName As String = "IApplicationDO_UpdateVerification"
        
    Dim objErrAssist As New ErrAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
    Dim objIDOAssist As IDOAssist

    If objContext Is Nothing Then
      Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadVerificationData()
    objIDOAssist.Update vxmlTableElement, xmlClassDefDoc
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Sub

UpdateVerificationVbErr:
        
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
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

End Sub

Private Sub IApplicationDO_CreateDeclaration(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   create an instance of the persistant data associated with this data object
'   for each set of data in the request
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo CreateDeclarationVbErr

    Const strFunctionName As String = "IApplicationDO_CreateDeclaration"
   
    Dim objErrAssist As New ErrAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
    Dim objIDOAssist As IDOAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
                    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadDeclarationData()
    objIDOAssist.Create vxmlTableElement, xmlClassDefDoc
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Sub

CreateDeclarationVbErr:
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    objErrAssist.AddToErrSource strFunctionName
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    If Not objContext Is Nothing Then
        If objErrAssist.IsSystemError Then
            objContext.SetAbort
        Else
            objContext.SetComplete
        End If
    End If
    
    Set objErrAssist = Nothing
    
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Sub

Private Sub IApplicationDO_DeleteAllDeclaration( _
    ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   Delete all instances of the persistant data associated with this
'   data object that match the key values specified
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo DeleteAllDeclarationVbErr

    Const strFunctionName As String = "IApplicationDO_DeleteAllDeclaration"

    Dim objErrAssist As New ErrAssist
    
    Dim objIDOAssist As IDOAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
                            
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
                
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadDeclarationData()
    
    Err.Clear
    On Error Resume Next
    objIDOAssist.DeleteAll vxmlTableElement, xmlClassDefDoc
    objErrAssist.SaveErr
    On Error GoTo DeleteAllDeclarationVbErr
    If objErrAssist.LastErr <> 0 Then
        If objErrAssist.GetOmigaErrorNumber(objErrAssist.LastErr) <> _
            oeNoRowsAffectedByDeleteAll Then
            objErrAssist.ReRaise
        End If
    End If
 
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Sub

DeleteAllDeclarationVbErr:
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    If Not objContext Is Nothing Then
        If objErrAssist.IsSystemError Then
            objContext.SetAbort
        Else
            objContext.SetComplete
        End If
    End If

    Set objErrAssist = Nothing

    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Private Function IApplicationDO_FindDeclarationList( _
    ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for all instances of the persistant data associated with
'   this data object for the values supplied
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo FindDeclarationListVbErr

    Const strFunctionName As String = "IApplicationDO_FindDeclarationList"

    Dim objErrAssist As New ErrAssist
    
    Dim objIDOAssist As IDOAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
    Dim xmlReturn As IXMLDOMNode
    Dim xmlAFFXML      As IXMLDOMNode, _
        xmlAFFResponse As IXMLDOMNode
    Dim strApplicationNumber         As String, _
        strApplicationFactFindNumber As String
    Dim objXmlAssist As XMLAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    Set objXmlAssist = New XMLAssist
         
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
     
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadDeclarationData()
    Set xmlReturn = objIDOAssist.FindList(vxmlTableElement, xmlClassDefDoc)

    ' Get ApplicationFactFind data (ostensibly for the BCDECLARATIONNOTES field)
    strApplicationNumber = objXmlAssist.GetTagValue(vxmlTableElement, "APPLICATIONNUMBER")
    strApplicationFactFindNumber = objXmlAssist.GetTagValue(vxmlTableElement, "APPLICATIONFACTFINDNUMBER")

    Set xmlAFFXML = vxmlTableElement.ownerDocument.createElement("APPLICATIONFACTFIND")
    xmlAFFXML.appendChild(xmlAFFXML.ownerDocument.createElement("APPLICATIONNUMBER")).Text = strApplicationNumber
    xmlAFFXML.appendChild(xmlAFFXML.ownerDocument.createElement("APPLICATIONFACTFINDNUMBER")).Text = strApplicationFactFindNumber

    Set xmlClassDefDoc = objIClassDef.LoadApplicationFactFindData()
    Set xmlAFFResponse = objIDOAssist.GetData(xmlAFFXML, xmlClassDefDoc)

    xmlReturn.appendChild xmlAFFResponse

    Set IApplicationDO_FindDeclarationList = xmlReturn

    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objErrAssist = Nothing
    Set objIDOAssist = Nothing
    Set xmlReturn = Nothing
    Set xmlAFFResponse = Nothing
    Set xmlAFFXML = Nothing
    Set objXmlAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

FindDeclarationListVbErr:
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set xmlReturn = Nothing
    Set xmlAFFResponse = Nothing
    Set xmlAFFXML = Nothing
    Set objXmlAssist = Nothing

    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    If Not objContext Is Nothing Then
        If objErrAssist.IsSystemError Then
            objContext.SetAbort
        Else
            objContext.SetComplete
        End If
    End If

    Set objErrAssist = Nothing
    
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Private Function IApplicationDO_FindOtherResidentDetailsList( _
    ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
' pass:
'   vxmlxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo FindOtherResidentDetailsListVbErr

    Const strFunctionName As String = "IApplicationDO_FindOtherResidentDetailsList"

    Dim objErrAssist As New ErrAssist
    
    Dim objIDOAssist As IDOAssist
    Dim objXmlAssist As New XMLAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIApplicationClassDef As IApplicationClassDef
    Dim xmlElement As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
         
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
     
    Set objIApplicationClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIApplicationClassDef.LoadOtherResidentDetailsData()
    
    Set xmlElement = objIDOAssist.FindListMultiple( _
        vxmlTableElement, xmlClassDefDoc, vstritemname:="OTHERRESIDENT")
       
    For Each xmlNode In xmlElement.selectNodes("OTHERRESIDENTDETAILS/OTHERRESIDENT")
        objXmlAssist.PromoteNode xmlElement, xmlNode
    Next xmlNode
       
    Set IApplicationDO_FindOtherResidentDetailsList = xmlElement.cloneNode(True)
    
    Set objIApplicationClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objErrAssist = Nothing
    Set objIDOAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlElement = Nothing
    Set xmlNode = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

FindOtherResidentDetailsListVbErr:
    
    Set objIApplicationClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlElement = Nothing
    Set xmlNode = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
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

Private Sub AddDateToRequest(ByRef robjXmlElem As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   Set the APPLICATIONDATE to the specified value if not already set.
'   If an APPLICATIONDATE tag does not exist, create one and set it
' pass:
'   robjXmlElem xml Application element
' return:       n/a
'------------------------------------------------------------------------------------------
On Error GoTo AddDateToRequestVbErr
    
    Const strFunctionName As String = "AddDateToRequest"
       
    Dim objXmlDateElem As IXMLDOMElement
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim strTagName As String
    
    strTagName = "APPLICATIONDATE"
    
    Set objXmlDateElem = _
        robjXmlElem.getElementsByTagName(strTagName).Item(0)
        
    ' Insert a tag if one isn't present
    If objXmlDateElem Is Nothing Then
        Set objXmlDateElem = _
            robjXmlElem.ownerDocument.createElement(strTagName)
        robjXmlElem.appendChild objXmlDateElem
    End If
    
    ' Set the value for the tag if there isn't a value already
    If objXmlDateElem.Text = "" Then
        objXmlDateElem.Text = Format(Now, "dd/mm/yyyy")
    End If
  
    Set objXmlDateElem = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
  
    Exit Sub

AddDateToRequestVbErr:
    
    Set objXmlDateElem = Nothing
    Set objXmlAssist = Nothing
    
    If objErrAssist.IsSystemError Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    Set objErrAssist = Nothing
    
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Sub

Private Sub AddCurrentDateTimeToRequest(ByVal vstrTagName As String, ByRef robjXmlElem As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   Adds the currentdate and time to the Request. Creates a new Tag, if required.
' pass:
'   vstrTagName the tagname to assign the DATETIME to
'   robjXmlElem xml Application element
' return:       n/a
'------------------------------------------------------------------------------------------
On Error GoTo AddCurrentDateTimeToRequestVbErr
    
    Const strFunctionName As String = "AddCurrentDateTimeToRequest"
       
    Dim objXmlDateElem As IXMLDOMElement
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim objSqlAssist As New SQLAssist
    
    Set objXmlDateElem = _
        robjXmlElem.getElementsByTagName(vstrTagName).Item(0)
        
    ' Insert a tag if one isn't present
    If objXmlDateElem Is Nothing Then
        Set objXmlDateElem = _
            robjXmlElem.ownerDocument.createElement(vstrTagName)
        robjXmlElem.appendChild objXmlDateElem
    End If
    
    ' Set the value for the tag if there isn't a value already
    objXmlDateElem.Text = Format(Now, "dd/mm/yyyy HH:MM:SS")
  
    Set objXmlDateElem = Nothing
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objSqlAssist = Nothing

    Exit Sub

AddCurrentDateTimeToRequestVbErr:
    
    Set objXmlDateElem = Nothing
    Set objXmlAssist = Nothing
    Set objSqlAssist = Nothing
    
    If objErrAssist.IsSystemError Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    Set objErrAssist = Nothing
    
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Sub

Private Sub IApplicationDO_CreateMemoPad(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   create an instance of the persistant data associated with this data object
'   for each set of data in the request
' pass:
'   vxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo CreateVbErr

    Const strFunctionName As String = "IApplicationDO_CreateMemoPad"
   
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
    
    Dim objIDOAssist As IDOAssist
    Dim objErrAssist As ErrAssist
    Set objErrAssist = New ErrAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
                    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
                           
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadMemoPadData()
    
    ' Add next MemoPadId
    objIDOAssist.GenerateSequenceNumber vxmlTableElement, xmlClassDefDoc, "MEMOPADID"
    
    objIDOAssist.Create vxmlTableElement, xmlClassDefDoc
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Sub

CreateVbErr:
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    If Not objContext Is Nothing Then
        If objErrAssist.IsSystemError Then
            objContext.SetAbort
        Else
            objContext.SetComplete
        End If
    End If
        
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Sub

Private Sub IApplicationDO_DeleteMemoPad(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   Delete a single instance of the persistant data associated with this
'   data object
' pass:
'   vxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo DeleteVbErr

    Const strFunctionName As String = "IApplicationDO_DeleteMemoPad"
    
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
    
    Dim objIDOAssist As IDOAssist
    Dim objErrAssist As ErrAssist
    Set objErrAssist = New ErrAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
                    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
                
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadMemoPadData()

    objIDOAssist.Delete vxmlTableElement, xmlClassDefDoc
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Sub

DeleteVbErr:

    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    If Not objContext Is Nothing Then
        If objErrAssist.IsSystemError Then
            objContext.SetAbort
        Else
            objContext.SetComplete
        End If
    End If
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Function IApplicationDO_FindMemoPadList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for all instances of the persistant data associated with
'   this data object for the values supplied
' pass:
'   vxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo FindListVbErr

    Const strFunctionName As String = "IApplicationDO_FindMemoPadList"
    
    Dim strSortOrder As String 'APWP1/BM067 - DPF 16/9/02
    Dim strField As String 'APWP1/BM067 - DPF 16/9/02
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
    Dim objIDOAssist As IDOAssist
    Dim objErrAssist As ErrAssist
    Set objErrAssist = New ErrAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
                    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If

    'APWP1/BM067 - DPF 16/09/02 - add sort to call for memo pad entries
    strSortOrder = m_objXmlAssist.GetNodeValue(vxmlTableElement, "//SORTORDER")
'GD 27/01/03 BM0275 START
'    Select Case strSortOrder
'        Case "10"
'            strField = "ENTRYDATETIME DESC"
'        Case "20"
'            strField = "USERID"
'        Case "30"
'            strField = "ENTRYTYPE"
'        Case "40"
'            strField = "SCREENREFERENCE"
'        Case Else
'            m_objErrAssist.ThrowError TypeName(Me), _
'                                      strFunctionName, _
'                                      oeInvalidParameter, _
'                                      "Sort order must be 10, 20, 30 or 40"
'    End Select
    strSortOrder = GetValidationTypeForValueID("MemoPadSortBy", CSafeInt(strSortOrder))
    Select Case strSortOrder
        Case "D"
            strField = "ENTRYDATETIME DESC"
        Case "U"
            strField = "USERID"
        Case "T"
            strField = "ENTRYTYPE"
        Case "S"
            strField = "SCREENREFERENCE"
        Case Else
            m_objErrAssist.ThrowError TypeName(Me), _
                                      strFunctionName, _
                                      oeInvalidParameter, _
                                      "Unrecognized sort order parameter for Memopad"
    End Select
    'GD 27/01/03 BM0275 END
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadMemoPadData()
    Set IApplicationDO_FindMemoPadList = objIDOAssist.FindList(vxmlTableElement, xmlClassDefDoc, strField)

    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objContext = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

FindListVbErr:
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objContext = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    If Not objContext Is Nothing Then
        If objErrAssist.IsSystemError Then
            objContext.SetAbort
        Else
            objContext.SetComplete
        End If
    End If
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Private Function IApplicationDO_GetMemoPadData(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   this data object
' pass:
'   vxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo GetDataVbErr

    Const strFunctionName As String = "IApplicationDO_GetMemoPadData"
      
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
    
    Dim objIDOAssist As IDOAssist
    Dim objErrAssist As ErrAssist
    Set objErrAssist = New ErrAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
                    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
   
    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadMemoPadData()
    Set IApplicationDO_GetMemoPadData = objIDOAssist.GetData(vxmlTableElement, xmlClassDefDoc)
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

GetDataVbErr:
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    If Not objContext Is Nothing Then
        If objErrAssist.IsSystemError Then
            objContext.SetAbort
        Else
            objContext.SetComplete
        End If
    End If
        
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Sub IApplicationDO_UpdateMemoPad(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   Update a single instance of the persistant data associated with this
'   data object
' pass:
'   vxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo UpdateVbErr

    Const strFunctionName As String = "IApplicationDO_UpdateMemoPad"
            
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IApplicationClassDef
    Dim objIDOAssist As IDOAssist
    
    Dim objErrAssist As ErrAssist
    Set objErrAssist = New ErrAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
                    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If

    Set objIClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIClassDef.LoadMemoPadData()
    objIDOAssist.Update vxmlTableElement, xmlClassDefDoc
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Sub

UpdateVbErr:
        
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    If Not objContext Is Nothing Then
        If objErrAssist.IsSystemError Then
            objContext.SetAbort
        Else
            objContext.SetComplete
        End If
    End If

    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Private Function IApplicationDO_GetMortgageAccountData(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   this data object
' pass:
'   vxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo GetMortgageAccountDataVbErr

    Const strFunctionName As String = "IApplicationDO_GetMortgageAccountData"
    
                            
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As ICFClassDef
    
    Set objIClassDef = New CFClassDef
    Set xmlClassDefDoc = objIClassDef.LoadMortgageAccountData()
    Set IApplicationDO_GetMortgageAccountData = m_objIDOAssist.GetData(vxmlTableElement, xmlClassDefDoc)
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    
    m_objContext.SetComplete
    
    
    Exit Function

GetMortgageAccountDataVbErr:
    
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    
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

Private Sub IApplicationDO_UpdateApplicationOffer(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   Update a single instance of the persistant data associated with this
'   data object
' pass:
'   vxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo UpdateApplicationOfferVbErr

    Const strFunctionName As String = "IApplicationDO_UpdateApplicationOffer"
    
    
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIApplicationClassDef As IApplicationClassDef
        
    Set objIApplicationClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIApplicationClassDef.LoadApplicationOfferData()
    m_objIDOAssist.Update vxmlTableElement, xmlClassDefDoc
    
    Set objIApplicationClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    
    m_objContext.SetComplete
        
    
    Exit Sub

UpdateApplicationOfferVbErr:
            
    
    Set objIApplicationClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    
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

End Sub

Private Sub IApplicationDO_CreateApplicationOffer(ByVal vxmlTableElement As IXMLDOMElement)
' header ----------------------------------------------------------------------------------
' description:
'   create an instance of the persistant data associated with this data object
'   for each set of data in the request
' pass:
'   vxmlTableElement  xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo CreateApplicationOfferVbErr

    Const strFunctionName As String = "IApplicationDO_CreateApplicationOffer"
    

    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIApplicationClassDef As IApplicationClassDef
                               
    Set objIApplicationClassDef = New ApplicationClassDef
    Set xmlClassDefDoc = objIApplicationClassDef.LoadApplicationOfferData()
    m_objIDOAssist.Create vxmlTableElement, xmlClassDefDoc
    
    Set objIApplicationClassDef = Nothing
    Set xmlClassDefDoc = Nothing

    m_objContext.SetComplete


    Exit Sub

CreateApplicationOfferVbErr:
    

    Set objIApplicationClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    
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

