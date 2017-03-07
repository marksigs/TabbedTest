VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "IntermediaryDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'Workfile:      IntermediaryDO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling

'Description:   Code for IntermediaryDO Data Object
'
'Dependencies:  IntermediaryBO
'
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'Prog   Date        Description
'SR     29/10/99    Created
'PSC    07/12/99    Amend to perform wildcard search on PanelId, Forename, Surname an Town
'LD     07/11/00    Explicity close recordsets
'LD     07/11/00    Explicity destroy command objects
'ASt    09/03/01    CC012: Now compatible with SQL Server and Oracle
'AW     01/10/01    SYS2518 Use stored procs in GetLeadAgentName
'JR     08/10/01    Omiplus24, Modified GetIndividualIntemediary, LoadContactDetails and added
'                   LoadContactTelephoneDetails
'AD     10/10/01    Removed the MoveFirst as it's a forward only recordset. SYS2718
'JR     26/10/01    SYS2680, comment out Omiplus24 telephone changes
'JR     03/12/01    SYS2459 re-apply Omiplus24 telephone changes
'SG     29/05/02    SYS4767 MSMS to Core integration
'------------------------------------------------------------------------------------------

Option Explicit


Public Function FindIndividualIntermediaryList(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:
'   Get the data the intermediary / intermediaries that match the criterion
'   as passed to the function
' pass:
'   vstrXMLRequest  xml Request data stream containing data to which identifies
'                   the creiteion to be used in the Query
' return:
'   FindIndividualIntermediaryList         string containing XML data stream representation of
'                                          data retrieved
' Raise Errors: if record not found, raise omiga4RecordNotFound
'------------------------------------------------------------------------------------------
On Error GoTo FindIndividualIntermediaryListVbErr

    Const strFunctionName As String = "FindIndividualIntermediaryList"
    
    Dim objXMLAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim objSQLAssist As New SQLAssist
     
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
            
    Dim strSql As String
    Dim strXML As String 'Output (XML string) of the function
          
    Dim objXmlIn As FreeThreadedDOMDocument40
    Dim objXmlOut As New FreeThreadedDOMDocument40
    objXmlOut.validateOnParse = False
    objXmlOut.setProperty "NewParser", True
    Dim objXmlDoc As New FreeThreadedDOMDocument40
    objXmlDoc.validateOnParse = False
    objXmlDoc.setProperty "NewParser", True
    Dim objXmlIntermedListElem As IXMLDOMElement
    Dim objXmlIndIntermedElem As IXMLDOMElement
        
    Dim objXmlNode As IXMLDOMNode
    
    Dim objADOAssist As ADOAssist
    Dim objDOAssist As DOAssist
        
    Set objXmlIn = objXMLAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    'Retrieve the values of the input params passed thru request string
    Dim strPanelId As String
    Dim strForename As String
    Dim strSurname As String
    Dim strFormattedValue As String
    Dim strTown As String
    Dim strCurrentDate As String
    Dim blnContainsWildcard As Boolean
   
    strPanelId = objXMLAssist.GetTagValue(objXmlIn.documentElement, "PANELID")
    strForename = objXMLAssist.GetTagValue(objXmlIn.documentElement, "FORENAME")
    strSurname = objXMLAssist.GetTagValue(objXmlIn.documentElement, "SURNAME")
    strTown = objXMLAssist.GetTagValue(objXmlIn.documentElement, "TOWN")
    
    
    If objContext Is Nothing Then
        Set objADOAssist = New ADOAssist
        Set objDOAssist = New DOAssist
    Else
        Set objADOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
        Set objDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
               
    'Get the current date and store it
    strCurrentDate = objSQLAssist.GetSystemDate
    
    'Build the SQL string to be executed
    strSql = "SELECT * FROM FINDINDIVIDUALINTERMEDIARYLIST WHERE " _
             & " INTERMEDIARYACTIVEFROM <= " & strCurrentDate _
             & " AND (INTERMEDIARYACTIVETO >= " & strCurrentDate & " OR INTERMEDIARYACTIVETO IS NULL) "
    
    If strPanelId <> "" Then
        strFormattedValue = objSQLAssist.FormatWildcardedString(strPanelId, _
                                                                blnContainsWildcard)
        If blnContainsWildcard = False Then
            strSql = strSql & " AND INTERMEDIARYPANELID = " & objSQLAssist.FormatString(strPanelId)
        Else
            strSql = strSql & " AND INTERMEDIARYPANELID " & _
                     objSQLAssist.GetLikeOperator() & " " & strFormattedValue
        End If

    End If
    
    If strForename <> "" Then
        strFormattedValue = objSQLAssist.FormatWildcardedString(strForename, _
                                                                blnContainsWildcard)
        If blnContainsWildcard = False Then
            strSql = strSql & " AND FORENAME = " & objSQLAssist.FormatString(strForename)
        Else
            strSql = strSql & " AND FORENAME " & _
                     objSQLAssist.GetLikeOperator() & " " & strFormattedValue
        End If
    End If
    
    If strSurname <> "" Then
        strFormattedValue = objSQLAssist.FormatWildcardedString(strSurname, _
                                                                blnContainsWildcard)
        If blnContainsWildcard = False Then
            strSql = strSql & " AND SURNAME = " & objSQLAssist.FormatString(strSurname)
        Else
            strSql = strSql & " AND SURNAME " & _
                     objSQLAssist.GetLikeOperator() & " " & strFormattedValue
        End If
    End If
    
    If strTown <> "" Then
        strFormattedValue = objSQLAssist.FormatWildcardedString(strTown, _
                                                                blnContainsWildcard)
        If blnContainsWildcard = False Then
            strSql = strSql & " AND TOWN = " & objSQLAssist.FormatString(strTown)
        Else
            strSql = strSql & " AND TOWN " & _
                     objSQLAssist.GetLikeOperator() & " " & strFormattedValue
        End If
    End If
          
    'Open the recordset with the SQL string built above
    Dim Cmd As New ADODB.Command
    Cmd.ActiveConnection = objADOAssist.GetConnStr
    Cmd.CommandText = strSql
    
    Dim rstThisRecSet As New ADODB.Recordset
    Set rstThisRecSet = Cmd.Execute(, , adCmdText)
    
    If rstThisRecSet.EOF And rstThisRecSet.BOF Then
        ' raise application error to be interpreted by calling object
        objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound
    End If
    
    Dim objAddressDO As New AddressDO
    
    'Build the templates (one for each table)  to build XML string (Output)
    Dim strIntermediaryDef As String, strIntIndividualDef As String, strAddressDef As String
    strIntermediaryDef = LoadIntermediaryList()
    strIntIndividualDef = LoadIntIndividualList()
    strAddressDef = objAddressDO.LoadAddressData
    
    Set objXmlIntermedListElem = objXmlOut.createElement("INDIVIDUALINTERMEDIARYLIST")
    objXmlOut.appendChild objXmlIntermedListElem
    
    'Build the output XML string
    'AD 10/10/01 Removed the MoveFirst as it's a forward only recordset. SYS2718
    
    ' Navigate through the record set and buid the string based on the templates defined above
    
    Do While Not rstThisRecSet.EOF
        Set objXmlIndIntermedElem = objXmlOut.createElement("INDIVIDUALINTERMEDIARY")
        objXmlIntermedListElem.appendChild objXmlIndIntermedElem
        'Add Intermediary XML string
        strXML = objDOAssist.GetXMLFromRecordSet(rstThisRecSet, strIntermediaryDef)
        Set objXmlDoc = objXMLAssist.load(strXML, TypeName(Me), strFunctionName)
        Set objXmlNode = objXmlIndIntermedElem.appendChild(objXmlDoc.documentElement)
        
        'Add Intermediary Individual XML string
        strXML = objDOAssist.GetXMLFromRecordSet(rstThisRecSet, strIntIndividualDef)
        Set objXmlDoc = objXMLAssist.load(strXML, TypeName(Me), strFunctionName)
        Set objXmlNode = objXmlIndIntermedElem.appendChild(objXmlDoc.documentElement)
        
        'Add Address XML string ONLY WHEN ADDRESSGUID IS NOT NULL
        
        If Not IsNull(rstThisRecSet.Fields("AddressGuid")) Then
            strXML = objDOAssist.GetXMLFromRecordSet(rstThisRecSet, strAddressDef)
            Set objXmlDoc = objXMLAssist.load(strXML, TypeName(Me), strFunctionName)
            Set objXmlNode = objXmlIndIntermedElem.appendChild(objXmlDoc.documentElement)
        End If
        rstThisRecSet.MoveNext
    Loop
    rstThisRecSet.Close

    FindIndividualIntermediaryList = objXmlOut.xml
    
    'Clear all the objects
    Set rstThisRecSet = Nothing
    Set Cmd = Nothing
    
    Set objDOAssist = Nothing
    Set objADOAssist = Nothing
    Set objAddressDO = Nothing
    
    Set objErrAssist = Nothing
    Set objXMLAssist = Nothing
    Set objSQLAssist = Nothing
    
    Set objXmlDoc = Nothing
    Set objXmlOut = Nothing
    Set objXmlIn = Nothing
    Set objXmlIntermedListElem = Nothing
    Set objXmlIndIntermedElem = Nothing
    Set objXmlNode = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

FindIndividualIntermediaryListVbErr:
    
    ' Close the recordset if still open
        If Not rstThisRecSet Is Nothing Then
                If rstThisRecSet.State = adStateOpen Then
                        rstThisRecSet.Close
                End If
        End If

    Set rstThisRecSet = Nothing
    Set Cmd = Nothing

    Set objADOAssist = Nothing
    Set objAddressDO = Nothing
        
    Set objXMLAssist = Nothing
    Set objSQLAssist = Nothing
    
    Set objXmlDoc = Nothing
    Set objXmlOut = Nothing
    Set objXmlIn = Nothing
    Set objXmlIntermedListElem = Nothing
    Set objXmlIndIntermedElem = Nothing
    Set objXmlNode = Nothing
        
    Set objDOAssist = Nothing
    
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

Public Function GetIndividualIntermediary(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:
'   Get the data the intermediary - from tables Intermediary, Individual Intermediary,
'   IntermediaryInsuranceType, CorrespondenceType, DailyReportDays, ContactDetails, Address
' pass:
'   vstrXMLRequest  xml Request data stream containing data to which identifies
'                   the instance of the persistant data to be retrieved
' return:
'   GetIndividualIntermediary         string containing XML data stream representation of
'                                      data retrieved
' Raise Errors: if record not found, raise omiga4RecordNotFound
'------------------------------------------------------------------------------------------
  On Error GoTo GetIndividualIntermediaryVbErr
        
    Const strFunctionName As String = "GetIndividualIntermediary"
    
    Dim objXMLAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim objSQLAssist As New SQLAssist
         
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim strSql As String 'Query to be executed
    Dim strXML As String 'Output (XML string) of the function
          
    Dim objXmlIn As FreeThreadedDOMDocument40
    Dim objXmlOut As New FreeThreadedDOMDocument40
    objXmlOut.validateOnParse = False
    objXmlOut.setProperty "NewParser", True
    Dim objXmlDoc As New FreeThreadedDOMDocument40
    objXmlDoc.validateOnParse = False
    objXmlDoc.setProperty "NewParser", True
        
    'Load the input string on a XML document and get the input parameter value
    Set objXmlIn = objXMLAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    Dim strIntermediaryGuid As String
    strIntermediaryGuid = objXMLAssist.GetTagValue(objXmlIn.documentElement, "INTERMEDIARYGUID")
    
    'Create DOAssist and ADOAssist objects to manipulate the data
    Dim objADOAssist As ADOAssist
    Dim objDOAssist As DOAssist
    If objContext Is Nothing Then
        Set objADOAssist = New ADOAssist
        Set objDOAssist = New DOAssist
    Else
        Set objADOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
        Set objDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    'class definition templates required for building the output XML string
    Dim strDefTemp As String
    Dim strIntermediaryDef As String
    
    Dim objXmlIndIntermediaryElem As IXMLDOMElement
    Dim objXmlIntInsurListElem As IXMLDOMElement
    Dim objXmlCorrTypeListElem As IXMLDOMElement
    Dim objXmlDailyRepDaysListElem As IXMLDOMElement
    Dim objXmlLeadAgentElem As IXMLDOMElement
    
    Dim objXmlNode As IXMLDOMNode

    Dim Cmd As New ADODB.Command
    Cmd.ActiveConnection = objADOAssist.GetConnStr
    Dim rstThisRecSet As New ADODB.Recordset
       
    Set objXmlIndIntermediaryElem = objXmlOut.createElement("INDIVIDUALINTERMEDIARY")
    objXmlOut.appendChild objXmlIndIntermediaryElem
    
    'Add the node for Intermediary table
    'Build the SQL string with input params passed and create recordset
    'ASt 09/03/01 CC012: Now compatible with SQL Server and Oracle
#If GENERIC_SQL Then
    strSql = "SELECT * FROM INTERMEDIARY WHERE INTERMEDIARYGUID = " _
        & objSQLAssist.FormatGuid(strIntermediaryGuid)
#Else
    strSql = "SELECT * FROM INTERMEDIARY WHERE INTERMEDIARYGUID = '" _
        & strIntermediaryGuid & "'"
#End If
    Cmd.CommandText = strSql
    Set rstThisRecSet = Cmd.Execute(, , adCmdText)
    
    If rstThisRecSet.EOF Then
        ' raise application error to be interpreted by calling object
        objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound
    End If
    
    'AD 10/10/01 Removed the MoveFirst as it's a forward only recordset. SYS2718
    
    Dim objAddressDO As New AddressDO
    
    'Add the details from recordset to the XML string
    strXML = objDOAssist.GetXMLFromRecordSet(rstThisRecSet, LoadIntermediaryList)
        rstThisRecSet.Close
    Set objXmlDoc = objXMLAssist.load(strXML, TypeName(Me), strFunctionName)
    Set objXmlNode = objXmlIndIntermediaryElem.appendChild(objXmlDoc.documentElement)
        
    'Add the node for IntermediaryIndividual table
    strDefTemp = LoadIntIndividualList()
#If GENERIC_SQL Then
    Call BuildAndAppendNodes("INTERMEDIARYINDIVIDUAL", strDefTemp, objXmlIndIntermediaryElem, _
                             "INTERMEDIARYGUID", objSQLAssist.FormatGuid(strIntermediaryGuid))
#Else
    Call BuildAndAppendNodes("INTERMEDIARYINDIVIDUAL", strDefTemp, objXmlIndIntermediaryElem, _
                             "INTERMEDIARYGUID", "'" & strIntermediaryGuid & "'")
#End If

    'Add node for Address table
    strDefTemp = objAddressDO.LoadAddressData
    Dim strTempCondition As String

#If GENERIC_SQL Then
    strTempCondition = "(SELECT ADDRESSGUID FROM INTERMEDIARY WHERE INTERMEDIARYGUID = " _
                         & objSQLAssist.FormatGuid(strIntermediaryGuid) & ")"
#Else
    strTempCondition = "(SELECT ADDRESSGUID FROM INTERMEDIARY WHERE INTERMEDIARYGUID = '" _
                         & strIntermediaryGuid & "')"
#End If

    Call BuildAndAppendNodes("ADDRESS", strDefTemp, objXmlIndIntermediaryElem, _
                             "ADDRESSGUID", strTempCondition)
                             
    'Add node for Contact Details table
    strDefTemp = LoadContactDetails
#If GENERIC_SQL Then
    strTempCondition = "(SELECT CONTACTDETAILSGUID FROM INTERMEDIARY WHERE INTERMEDIARYGUID = " _
                         & objSQLAssist.FormatGuid(strIntermediaryGuid) & ")"
#Else
    strTempCondition = "(SELECT CONTACTDETAILSGUID FROM INTERMEDIARY WHERE INTERMEDIARYGUID = '" _
                         & strIntermediaryGuid & "')"
#End If
    Call BuildAndAppendNodes("CONTACTDETAILS", strDefTemp, objXmlIndIntermediaryElem, _
                             "CONTACTDETAILSGUID", strTempCondition)
       
    'JR - Omiplus, append Node for ContactTelephoneDetails table to ContactDetails
    Dim xmlContactDetailsNode As IXMLDOMNode
    Set xmlContactDetailsNode = objXmlIndIntermediaryElem.selectSingleNode("CONTACTDETAILS")
    If Not xmlContactDetailsNode Is Nothing Then
        objXmlIndIntermediaryElem.appendChild xmlContactDetailsNode
        strDefTemp = LoadContactTelephoneDetails
        Call BuildAndAppendNodes("CONTACTTELEPHONEDETAILS", strDefTemp, xmlContactDetailsNode, _
                             "CONTACTDETAILSGUID", strTempCondition)
    End If
    
    'Add List element for IntermediaryInsuranceType and nodes to it from that table
    Set objXmlIntInsurListElem = objXmlOut.createElement("INTERMEDIARYINSURANCETYPELIST")
    objXmlIndIntermediaryElem.appendChild objXmlIntInsurListElem
    
    strDefTemp = LoadIntermediaryInsuranceType()
#If GENERIC_SQL Then
    Call BuildAndAppendNodes("INTERMEDIARYINSURANCETYPE", strDefTemp, objXmlIntInsurListElem, _
                             "INTERMEDIARYGUID", objSQLAssist.FormatGuid(strIntermediaryGuid))
#Else
    Call BuildAndAppendNodes("INTERMEDIARYINSURANCETYPE", strDefTemp, objXmlIntInsurListElem, _
                             "INTERMEDIARYGUID", "'" & strIntermediaryGuid & "'")
#End If

    'Add List element for CorrespondenceType and nodes to it from that table
    Set objXmlCorrTypeListElem = objXmlOut.createElement("CORRESPONDENCETYPELIST")
    objXmlIndIntermediaryElem.appendChild objXmlCorrTypeListElem
    
    strDefTemp = LoadCorrespondenceType()
#If GENERIC_SQL Then
    Call BuildAndAppendNodes("CORRESPONDENCETYPE", strDefTemp, objXmlCorrTypeListElem, _
                             "INTERMEDIARYGUID", objSQLAssist.FormatGuid(strIntermediaryGuid))
#Else
    Call BuildAndAppendNodes("CORRESPONDENCETYPE", strDefTemp, objXmlCorrTypeListElem, _
                             "INTERMEDIARYGUID", "'" & strIntermediaryGuid & "'")
#End If
        
    'Add List element for DailyReportDays and nodes from that table
    Set objXmlDailyRepDaysListElem = objXmlOut.createElement("DAILYREPORTSLIST")
    objXmlIndIntermediaryElem.appendChild objXmlDailyRepDaysListElem
    
    strDefTemp = LoadDailyReportDays
#If GENERIC_SQL Then
    Call BuildAndAppendNodes("DAILYREPORTDAYS", strDefTemp, objXmlDailyRepDaysListElem, _
                             "INTERMEDIARYGUID", objSQLAssist.FormatGuid(strIntermediaryGuid))
#Else
    Call BuildAndAppendNodes("DAILYREPORTDAYS", strDefTemp, objXmlDailyRepDaysListElem, _
                             "INTERMEDIARYGUID", "'" & strIntermediaryGuid & "'")
#End If

    'Build input string and get the lead agent's name.
    strXML = "<INTERMEDIARYGUID>"
    strXML = strXML & strIntermediaryGuid
    strXML = strXML & "</INTERMEDIARYGUID>"
    
    strXML = GetLeadAgentName(strXML)
    
    If strXML = "" Then
        Set objXmlLeadAgentElem = objXmlOut.createElement("LEADAGENTNAME")
        objXmlIndIntermediaryElem.appendChild objXmlLeadAgentElem
    Else
        Set objXmlDoc = objXMLAssist.load(strXML, TypeName(Me), strFunctionName)
        Set objXmlNode = objXmlIndIntermediaryElem.appendChild(objXmlDoc.documentElement)
    End If
        
    'Assign the Output String
    GetIndividualIntermediary = objXmlOut.xml
    
    'Clear all the objects
    Set rstThisRecSet = Nothing
    Set Cmd = Nothing
    
    Set objDOAssist = Nothing
    Set objADOAssist = Nothing
    Set objAddressDO = Nothing
    
    Set objErrAssist = Nothing
    Set objXMLAssist = Nothing
    
    Set objXmlDoc = Nothing
    Set objXmlOut = Nothing
    Set objXmlIn = Nothing
    
    Set objXmlIndIntermediaryElem = Nothing
    Set objXmlIntInsurListElem = Nothing
    Set objXmlCorrTypeListElem = Nothing
    Set objXmlDailyRepDaysListElem = Nothing
    Set objXmlLeadAgentElem = Nothing
    Set xmlContactDetailsNode = Nothing 'JR - Omiplus24
    
    Set objXmlNode = Nothing
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

GetIndividualIntermediaryVbErr:
    
        ' Close the recordset if still open
        If Not rstThisRecSet Is Nothing Then
                If rstThisRecSet.State = adStateOpen Then
                        rstThisRecSet.Close
                End If
        End If

        Set rstThisRecSet = Nothing
    Set Cmd = Nothing
        Set objADOAssist = Nothing
    Set objAddressDO = Nothing
      
    Set objXMLAssist = Nothing
    Set objXmlDoc = Nothing
    Set objXmlOut = Nothing
    Set objXmlIn = Nothing
    
    Set objXmlIndIntermediaryElem = Nothing
    Set objXmlIntInsurListElem = Nothing
    Set objXmlCorrTypeListElem = Nothing
    Set objXmlDailyRepDaysListElem = Nothing
    Set objXmlLeadAgentElem = Nothing
    Set xmlContactDetailsNode = Nothing 'JR - Omiplus24
    
    Set objXmlNode = Nothing
    
    Set objDOAssist = Nothing
    
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
    
    're-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function


Private Function LoadIntermediaryList()
    Dim strXML As String
    
    strXML = strXML & "<TABLENAME>"
    strXML = strXML & "INTERMEDIARY"
    strXML = strXML & "<PRIMARYKEY>INTERMEDIARYGUID<TYPE>dbdtGuid</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>ADDRESSGUID<TYPE>dbdtGuid</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>CONTACTDETAILSGUID<TYPE>dbdtGuid</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>ISATTACHEDTOINTERMEDIARYGUID<TYPE>dbdtGuid</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>INTERMEDIARYACTIVEFROM<TYPE>dbdtDate</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>INTERMEDIARYACTIVETO<TYPE>dbdtDate</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>COMMUNICATIONMETHOD<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "<COMBO>CommunicationMethod</COMBO></OTHERS>"
    strXML = strXML & "<OTHERS>INTERMEDIARYPANELID<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>PROCURATIONFEELTVLIMIT<TYPE>dbdtDouble</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>PROCURATIONFEEPAYMENTMETHOD<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "<COMBO>ProcurationFeePaymentMethod</COMBO></OTHERS>"
    strXML = strXML & "<OTHERS>REPORTFREQUENCY<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "<COMBO>IntermediaryReportFrequency</COMBO></OTHERS>"
    strXML = strXML & "</TABLENAME>"

    LoadIntermediaryList = strXML
    
End Function

Private Function LoadIntIndividualList()
    Dim strXML As String
    
    strXML = strXML & "<TABLENAME>"
    strXML = strXML & "INTERMEDIARYINDIVIDUAL"
    strXML = strXML & "<PRIMARYKEY>INTERMEDIARYGUID<TYPE>dbdtGuid</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>FORENAME<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>SURNAME<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>TITLE<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "<COMBO>Title</COMBO></OTHERS>"
    strXML = strXML & "<OTHERS>USERROLEUSERID<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>USERROLEUNITID<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>KEYDEVELOPMENTAREA<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "<COMBO>KeyDevelopmentArea</COMBO></OTHERS>"
    
    'SG 29/05/02 SYS4767 START
    'STB: MSMS0011 - IntermediaryStatus.
    strXML = strXML & "<OTHERS>INTERMEDIARYSTATUS<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "<COMBO>IntermediaryStatus</COMBO></OTHERS>"
    'STB: MSMS0011 - End.
    'SG 29/05/02 SYS4767 END
    
    
    strXML = strXML & "</TABLENAME>"
    
    LoadIntIndividualList = strXML
End Function


'This builds the template required for getting the data from the table ContactDetails
'in the form of XML string
Private Function LoadContactDetails()
    Dim strXML As String
    
    strXML = strXML & "<TABLENAME>"
    strXML = strXML & "CONTACTDETAILS"
    strXML = strXML & "<PRIMARYKEY>CONTACTDETAILSGUID<TYPE>dbdtGuid</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>CONTACTFORENAME<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>CONTACTSURNAME<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>CONTACTTITLE<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>CONTACTTYPE<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "<COMBO>ContactType</COMBO></OTHERS>"
    strXML = strXML & "<OTHERS>EMAILADDRESS<TYPE>dbdtString</TYPE></OTHERS>"
 '   JR - Omiplus24, removed as part of telephone number change
'    strXML = strXML & "<OTHERS>FAXNUMBER<TYPE>dbdtString</TYPE></OTHERS>"
'    strXML = strXML & "<OTHERS>TELEPHONEEXTENSIONNUMBER<TYPE>dbdtString</TYPE></OTHERS>"
'    strXML = strXML & "<OTHERS>TELEPHONENUMBER<TYPE>dbdtString</TYPE></OTHERS>"
    
    strXML = strXML & "</TABLENAME>"
    
    LoadContactDetails = strXML
End Function


'This builds the template required for getting the data from the table DAILYREPORTDAYS
'in the form of XML string
Private Function LoadDailyReportDays()
    Dim strXML As String
    
    strXML = strXML & "<TABLENAME>"
    strXML = strXML & "DAILYREPORTDAYS"
    strXML = strXML & "<PRIMARYKEY>INTERMEDIARYGUID<TYPE>dbdtGuid</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>DAY<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"
    
    LoadDailyReportDays = strXML
End Function
    
'This builds the template required for getting the data from the table CORRESPONDENCETYPE
'in the form of XML string
    
Private Function LoadCorrespondenceType()
    Dim strXML As String

    strXML = strXML & "<TABLENAME>"
    strXML = strXML & "CORRESPONDENCETYPE"
    strXML = strXML & "<PRIMARYKEY>INTERMEDIARYGUID<TYPE>dbdtGuid</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>CORRESPONDTYPE<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "<COMBO>IntermediaryCorrespondenceType</COMBO></OTHERS>"
    strXML = strXML & "</TABLENAME>"
    
    LoadCorrespondenceType = strXML
End Function
    
'This builds the template required for getting the data from the table INTERMEDIARYINSURANCETYPE
'in the form of XML string
Private Function LoadIntermediaryInsuranceType()
    Dim strXML As String

    strXML = strXML & "<TABLENAME>"
    strXML = strXML & "INTERMEDIARYINSURANCETYPE"
    strXML = strXML & "<PRIMARYKEY>INTERMEDIARYGUID<TYPE>dbdtGuid</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>INSURANCETYPE<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "<COMBO>IntermediaryInsuranceType</COMBO></OTHERS>"
    strXML = strXML & "</TABLENAME>"
    
    LoadIntermediaryInsuranceType = strXML
End Function

Private Function BuildAndAppendNodes(ByVal strTable As String, _
                                 ByVal strClassDef As String, _
                                 ByVal objParentXMLElem As IXMLDOMElement, _
                                 ByVal strLSCondition As String, _
                                 ByVal strRSCondition As String) As String
'*Header*********************************************************************
'Description: Builds the node from the table from the records that match the criteria
'              and appends them to the parent element
'Pass : TableName, Class Definition Template, Lesft side of the contion, Right side of the
'       condtion to be applied on the table
'Return: XML String
'****************************************************************************
On Error GoTo BuildAndAppendNodesVbErr
        
    Const strFunctionName As String = "BuildAndAppendNodes"
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim objADOAssist As ADOAssist
    Dim objDOAssist As DOAssist
    Dim objErrAssist As New ErrAssist
    Dim objSQLAssist As New SQLAssist
    Dim objXMLAssist As New XMLAssist
    
    Dim strSql As String, strXML As String
    Dim Cmd As New ADODB.Command
    Dim rstThisRecSet As New ADODB.Recordset
    
    Dim objXmlDoc As New FreeThreadedDOMDocument40
    objXmlDoc.validateOnParse = False
    objXmlDoc.setProperty "NewParser", True
    Dim objXmlNode As IXMLDOMNode
    
    If objContext Is Nothing Then
        Set objADOAssist = New ADOAssist
        Set objDOAssist = New DOAssist
    Else
        Set objADOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
        Set objDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    strSql = "SELECT * FROM " & strTable & " WHERE " & strLSCondition & " = " & strRSCondition
    Cmd.ActiveConnection = objADOAssist.GetConnStr
    Cmd.CommandText = strSql
    Set rstThisRecSet = Cmd.Execute(, , adCmdText)
     
    If rstThisRecSet.EOF Then 'if no matching record exists return empty string
        strXML = ""
    Else
        'AD 10/10/01 Removed the MoveFirst as it's a forward only recordset. SYS2718
        Do While Not rstThisRecSet.EOF
            'Add the details from the recordset to XML string
            
            strXML = objDOAssist.GetXMLFromRecordSet(rstThisRecSet, strClassDef)
            Set objXmlDoc = objXMLAssist.load(strXML, TypeName(Me), strFunctionName)
            Set objXmlNode = objParentXMLElem.appendChild(objXmlDoc.documentElement)
            rstThisRecSet.MoveNext
        Loop
    End If
    rstThisRecSet.Close

    BuildAndAppendNodes = strXML
    
    Set objADOAssist = Nothing
    Set objDOAssist = Nothing
    Set objXMLAssist = Nothing
    Set objErrAssist = Nothing
    Set rstThisRecSet = Nothing
    Set objXmlDoc = Nothing
    Set objXmlNode = Nothing
    Set Cmd = Nothing
    
    Exit Function
BuildAndAppendNodesVbErr:

    ' Close the recordset if still open
        If Not rstThisRecSet Is Nothing Then
                If rstThisRecSet.State = adStateOpen Then
                        rstThisRecSet.Close
                End If
        End If

    Set objADOAssist = Nothing
    Set objDOAssist = Nothing
    Set objXMLAssist = Nothing
    Set rstThisRecSet = Nothing
    Set objXmlDoc = Nothing
    Set objXmlNode = Nothing
    
    Set Cmd = Nothing
    
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
    
    're-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function GetLeadAgentName(ByVal vstrXMLRequest As String) As String
'*Header*******************************************************************************
'Description: Gets the lead agent name for the current Intermediary
'Pass : XML containing IsAttachedtoIntermediaryGuid of the current Intermediary
'Return: XML String with lead agent name. Empty string if there is no lead agent
'*****************************************************************************************
On Error GoTo GetLeadAgentNameVbErr
        
    Const strFunctionName As String = "GetLeadAgentName"
    
    Dim adoConnection As New ADODB.Connection
    Dim adoCommand As New ADODB.Command
    Dim adoParameter As ADODB.Parameter
    
    Dim objXMLAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim objSQLAssist As New SQLAssist
    
    Dim objDOAssist As DOAssist
    Dim objADOAssist As ADOAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim strXmlOut As String 'XML string returned by the function
    Dim strSql As String 'SQL string to be executed
    
    If objContext Is Nothing Then
        Set objADOAssist = New ADOAssist
        Set objDOAssist = New DOAssist
    Else
        Set objADOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
        Set objDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
        
    'Find the value of IntermediaryGuid
    Set xmlIn = objXMLAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    Dim strIntermediaryGuid As String
    strIntermediaryGuid = objXMLAssist.GetTagValue(xmlIn.documentElement, "INTERMEDIARYGUID")
    
    Dim rstThisRecSet As New ADODB.Recordset
    
    'Check whether this is attached to any organisation
#If GENERIC_SQL Then
    strSql = "Select IsAttachedToIntermediaryGuid From Intermediary" _
             & " Where IntermediaryGuid = " & objSQLAssist.FormatGuid(strIntermediaryGuid)
#Else
    strSql = "Select IsAttachedToIntermediaryGuid From Intermediary" _
             & " Where IntermediaryGuid = " & objSQLAssist.FormatString(strIntermediaryGuid)
#End If
    rstThisRecSet.Open strSql, objADOAssist.GetConnStr, adOpenStatic
    
    If IsNull(rstThisRecSet!IsAttachedToIntermediaryGuid) Then
        GetLeadAgentName = ""
        Exit Function
    End If
    
    'Build the SQL string to be executed
' AW    01/10/01    SYS2518
#If GENERIC_SQL Then
    strSql = "USP_GETINTMDYLINKEDTOORG"
    Set adoParameter = adoCommand.CreateParameter("INTERMEDIARYGUID", adBinary, adParamInput, 16)
#Else
    strSql = "sp_GetIntmdyLinkedToOrg.GetLeadIntermediary"
    Set adoParameter = adoCommand.CreateParameter("INTERMEDIARYGUID", adBSTR, adParamInput)
#End If
    
    adoConnection.ConnectionString = objADOAssist.GetConnStr
    adoConnection.Open
    Set adoCommand.ActiveConnection = adoConnection
    adoCommand.CommandType = adCmdStoredProc
    adoCommand.CommandText = strSql

    adoParameter.Attributes = adParamNullable
    adoCommand.Parameters.Append adoParameter

    If Len(strIntermediaryGuid) > 0 Then
#If GENERIC_SQL Then
        adoCommand(0) = objSQLAssist.GuidStringToByteArray(strIntermediaryGuid)
#Else
        adoCommand(0) = strIntermediaryGuid
#End If
    Else
        adoCommand(0) = Null
    End If
    
    'Open the recordset with the SQL string built above
    rstThisRecSet.Close
    
    Set rstThisRecSet = adoCommand.Execute
    'rstThisRecSet.Open strSql, objADOAssist.GetConnStr, adOpenStatic, , adCmdText
   
    If rstThisRecSet.EOF And rstThisRecSet.BOF Then
        ' raise application error to be interpreted by calling object
        objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound, _
                                   "Lead Agent Guid "
    End If
    'Get the Intermediary GUID of the lead agent (last row in the recordset)
    'rstThisRecSet.MoveLast
    Dim strLeadIntermediaryGuid As String
    strLeadIntermediaryGuid = objSQLAssist.GuidToString(rstThisRecSet.Fields.Item("IntermediaryLinkedToOrg").Value)
    
    'Get the name of the Lead Intermediary
    'Build the SQL string to be executed open the recordset
    Dim Cmd As New ADODB.Command
#If GENERIC_SQL Then
    strSql = "SELECT name from IntermediaryOrganisation WHERE IntermediaryGuid = " _
             & objSQLAssist.FormatGuid(strLeadIntermediaryGuid)
#Else
    strSql = "SELECT name from IntermediaryOrganisation WHERE IntermediaryGuid = " _
             & objSQLAssist.FormatString(strLeadIntermediaryGuid)
#End If
    Cmd.ActiveConnection = objADOAssist.GetConnStr
    Cmd.CommandText = strSql
    rstThisRecSet.Close
    Set rstThisRecSet = Cmd.Execute(, , adCmdText)
    If rstThisRecSet.EOF And rstThisRecSet.BOF Then
        ' raise application error to be interpreted by calling object
        objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound, _
                                   "Lead Agent Organisation Name "
    End If
    
    Dim strLeadAgentName As String
    strLeadAgentName = rstThisRecSet!Name
        rstThisRecSet.Close
    
    strXmlOut = "<LEADAGENTNAME>"
    strXmlOut = strXmlOut & strLeadAgentName
    strXmlOut = strXmlOut & "</LEADAGENTNAME>"
    
    GetLeadAgentName = strXmlOut
        
    'Clear all the objects
    Set adoCommand.ActiveConnection = Nothing
    Set rstThisRecSet = Nothing
    Set Cmd = Nothing

    Set objErrAssist = Nothing
    Set objXMLAssist = Nothing
    Set objSQLAssist = Nothing

    Set objDOAssist = Nothing
    Set objADOAssist = Nothing
    
    Set xmlIn = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

GetLeadAgentNameVbErr:
    
    ' Close the recordset if still open
        If Not rstThisRecSet Is Nothing Then
                If rstThisRecSet.State = adStateOpen Then
                        rstThisRecSet.Close
                End If
        End If

    Set rstThisRecSet = Nothing
    Set Cmd = Nothing

    Set objErrAssist = Nothing
    Set objXMLAssist = Nothing
    Set objSQLAssist = Nothing

    Set objDOAssist = Nothing
    Set objADOAssist = Nothing
    
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

'JR - Omiplus24
'This builds the template required for getting the data from the table CONTACTTELEPHONEDETAILS
'in the form of XML string
Private Function LoadContactTelephoneDetails()
    Dim strXML As String
    
    strXML = strXML & "<TABLENAME>"
    strXML = strXML & "CONTACTTELEPHONEDETAILS"
    strXML = strXML & "<PRIMARYKEY>CONTACTDETAILSGUID<TYPE>dbdtGuid</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>TELEPHONESEQNUM<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>USAGE<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "<COMBO>ContactTelephoneUsage</COMBO></OTHERS>"
    strXML = strXML & "<OTHERS>COUNTRYCODE<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>AREACODE<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>TELENUMBER<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>EXTENSIONNUMBER<TYPE>dbdtString</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"
    
    LoadContactTelephoneDetails = strXML

End Function
