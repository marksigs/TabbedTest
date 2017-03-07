VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "RiskAssessmentDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      Omiga4DO.cls
'Copyright:     Copyright © 2000 Marlborough Stirling
'
'Description:   DO for Risk Assessment management. Used to retrieve all of the
'  Risk rules for a particular application at its current stage and then allow
'  authorised users the opportunity to override the risk assessment and lend the money
'  anyway.
'Dependencies:  Add any other dependent components
'
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'MH     29/03/00    Created
'MH     31/05/00    StoreRiskAssResults now returns a success XML string. This
'                   String identifies which record was created. This information can
'                   then be used by the TXBO and update the Underwriter's decision
'                   based on the internal business rules stored in the SPM.
'IVW    27/09/00    SYS1587 - Included SQLAssist, does not use MTS for SQLAssist either
'APS    03/10/00    Added back in SQLAssist object created as 'new'
'LD     07/11/00    Explicity close database connections
'LD     07/11/00    Explicity close recordsets
'PSC    06/03/01    SYS1879 Add GetLatestRiskAssesmentScore
'APS    19/03/01    SYS2074
'DM     16/05/01    SYS2314
'LD     11/06/01    SYS2367 SQL Server Port - Length must be specified in calls to CreateParameter
'DM     20/07/01    SYS2511 Changed datatype adVarChar
'CL     07/04/02    SYS4510  Modifications to class initialise
'CL     10/05/02    SYS4510  Remove class initialize & class terminate
'------------------------------------------------------------------------------------------

Option Explicit

Implements IRiskAssessmentDO
Implements ObjectControl
 
'=============================================
'Constant Declaration Section
'=============================================

'=============================================
'Variable Declaration Section
'=============================================

Private m_objContext As ObjectContext
Private m_objErrAssist As ErrAssist
Private m_objXmlAssist As XMLAssist
Private m_objSQLAssist As SQLAssist
Private m_objADOAssist As ADOAssist
Private m_objIDOAssist As IDOAssist
Private m_objIClassDef As IClassDefomRA



'------------------------------------------------------------------------------------------
'BMIDS Specific History:
'
' Prog  Date        Description
' MV    17/06/2002  BMIDS00032 - BM078- Amended IRiskAssessmentDO_GetRiskAssessment with ScoreCardInd
' MV    19/07/2002  BMIDS00219 -  Modified CreateRiskAssessmentRule - Amended ScoreCardInd
' DB    29/10/2002  BMIDS00712 - CC019 - Amended GetPreviousOverrideReasons
' DM    20/11/2002  BMIDS01022  Added NOLOCK hint to GetRiskAssessmentSeq
' DB    25/11/2002  BMIDS01065 - CC019 Further amendment to GetPreviousOverrideReasons
' MV    29/01/2003  BM0302  Amended IRiskAssessmentDO_GetPreviousOverrideReasons
' LD    03/02/2003  BM0397  Improve query in RiskAssessmentDO.IRiskAssessmentDO_GetRAApplicationStages
' SR    09/08/2004  BMIDS815 - new method CheckRiskAssessmentExistsForApplication
'------------------------------------------------------------------------------------------
'MARS Specific History:
' HMA   19/12/2005  MAR550 Allow for no Stage Number in GetRiskAssessment
' DRC   10/02/2006  MAR1209 - remove stage from GetRiskAssessment calls
' HMA   21/02/2006  MAR1291 - put stage back in.
' HMA   23/02/2006  MAR1315 - Change GetRAApplicationStages
' PE    05/04/2006  MAR1563 - Return total number of cases
'------------------------------------------------------------------------------------------

Private Sub ObjectControl_Activate()

    Set m_objContext = GetObjectContext()
    Set m_objErrAssist = New ErrAssist
    Set m_objXmlAssist = New XMLAssist
    Set m_objSQLAssist = New SQLAssist
    
    If m_objContext Is Nothing Then
        Set m_objADOAssist = New ADOAssist
        Set m_objIDOAssist = New DOAssist
        Set m_objIClassDef = New omRAClassDef
    Else
        Set m_objADOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
        Set m_objIDOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
        Set m_objIClassDef = m_objContext.CreateInstance(gstrRISK_ASSESSMENT & ".omRAClassDef")
    End If

End Sub

Private Function ObjectControl_CanBePooled() As Boolean

    ObjectControl_CanBePooled = False

End Function

Private Sub ObjectControl_Deactivate()

    Set m_objContext = Nothing
    Set m_objErrAssist = Nothing
    Set m_objXmlAssist = Nothing
    Set m_objSQLAssist = Nothing
    Set m_objADOAssist = Nothing
    Set m_objIDOAssist = Nothing
    Set m_objIClassDef = Nothing
End Sub

Private Function IRiskAssessmentDO_GetLatestRiskAssessment(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
On Error GoTo GetDataVbErr
    Const strFunctionName  As String = "GetLatestRiskAssessment"
    
    Set IRiskAssessmentDO_GetLatestRiskAssessment = IRiskAssessmentDO_GetRiskAssessment(vxmlTableElement)
    Exit Function
    
GetDataVbErr:
    
    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
           
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Function IRiskAssessmentDO_GetPreviousOverrideReasons(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
    
    On Error GoTo GetDataVbErr
    Const strFunctionName  As String = "GetPreviousOverrideReasons"

    Dim objIClassDef As IClassDefomRA
    Dim comSQL As ADODB.Command
    Dim objADOAssist As ADOAssist
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponse As IXMLDOMElement
    Dim xmlText As IXMLDOMElement
    Dim xmlRAIn As IXMLDOMElement
    
    Dim strApplicationNumber As String
    Dim strAppFactFindNo As String
    Dim strStageNumber As String
    Dim strRASequenceNumber As String
    Dim strRARuleNumber As String
    Dim strOut As String
    
    Dim rstRA As Recordset
    
    'DB BMIDS00712 - Added
    Dim objIDOAssist As IDOAssist, _
        objIADOAssist As IADOAssist, _
        objXmlAssist As New XMLAssist, _
        objSQLAssist As New SQLAssist
    Dim xmlDoc As New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    Dim strSQL As String
    Dim cmd As ADODB.Command, _
        rst As ADODB.Recordset
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    'DB End
    
    'Retrieve the minimum required input parameters
    If vxmlTableElement.nodeName <> "RISKASSESSMENTRULEOVERRIDE" Then
        Set xmlRAIn = m_objXmlAssist.GetMandatoryNode(vxmlTableElement, "RISKASSESSMENTRULEOVERRIDE")
    Else
        Set xmlRAIn = vxmlTableElement
    End If
    
    strApplicationNumber = m_objXmlAssist.GetMandatoryElementText(xmlRAIn, "APPLICATIONNUMBER")
    strAppFactFindNo = m_objXmlAssist.GetMandatoryElementText(xmlRAIn, "APPLICATIONFACTFINDNUMBER")
    strStageNumber = m_objXmlAssist.GetMandatoryElementText(xmlRAIn, "STAGENUMBER")
    strRASequenceNumber = m_objXmlAssist.GetElementText(xmlRAIn, "RISKASSESSMENTSEQUENCENUMBER")
    strRARuleNumber = m_objXmlAssist.GetElementText(xmlRAIn, "RARULENUMBER")
    
    'Create an output document and assign it
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponse = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponse
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
         
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
        Set objIADOAssist = New ADOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
        Set objIADOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    End If
    
'DB BMIDS00712 - Amended this to create a proper xml output document which can then be manipulated in the GUI.
    'Run the command
    'If RASeqeuencenumber is missing then get latest otherwise get the stated one
'    Set comSQL = New ADODB.Command
'    With comSQL
'        .ActiveConnection = m_objADOAssist.GetConnStr
'        .CommandType = adCmdText
'        .CommandText = "Select RAOVERRIDEDATETIME, USERID, RAOVERRIDEREASON, RAOVERRIDEAPPROVALREFERENCE" & _
'                " From RISKASSESSMENTRULEOVERRIDE where APPLICATIONNUMBER='" & strApplicationNumber & "'" & _
'                " AND " & _
'                "( " & _
'                    "( APPLICATIONFACTFINDNUMBER=" & strAppFactFindNo & " AND STAGENUMBER=" & strStageNumber & " AND RISKASSESSMENTSEQUENCENUMBER<" & strRASequenceNumber & ")" & _
'                " OR (APPLICATIONFACTFINDNUMBER<" & strAppFactFindNo & ") )" & _
'                " AND " & _
'                "RARULENUMBER=" & strRARuleNumber & _
'                " ORDER BY APPLICATIONFACTFINDNUMBER DESC, STAGENUMBER DESC, RISKASSESSMENTSEQUENCENUMBER  DESC"
'
'        Set rstRA = .Execute
'    End With
'
'    strOut = vbNullString
'
'    'Create a summary string if there is any data
'    If Not rstRA.BOF Then
'        Do While Not rstRA.EOF
'             strOut = strOut & FormatOverrideReason(rstRA, vblnWithRef:=True) & vbCrLf
'            rstRA.MoveNext
'        Loop
'    End If
'    rstRA.Close
'        Set rstRA = Nothing
'        Set comSQL = Nothing

'DB BMIDS00712 - END
    
'DB BMIDS00712 - New code to get records from DB and create proper xml output doc.
'DB BMIDS01065 - Changed sql so that all raoveridereasons are returned instead of just ones that
'                have a sequence number which is less than the current overriden rule !
    strSQL = "Select RAOVERRIDEDATETIME, USERID, RAOVERRIDEREASON  ,RAOVERRIDEREASONCODE, RAOVERRIDEAPPROVALREFERENCE" & _
                " From RISKASSESSMENTRULEOVERRIDE where APPLICATIONNUMBER='" & strApplicationNumber & "'" & _
                " AND " & _
                "( " & _
                    "( APPLICATIONFACTFINDNUMBER=" & strAppFactFindNo & " AND STAGENUMBER=" & strStageNumber & " AND RISKASSESSMENTSEQUENCENUMBER<=" & strRASequenceNumber & ")" & _
                " OR (APPLICATIONFACTFINDNUMBER<" & strAppFactFindNo & ") )" & _
                " AND " & _
                "RARULENUMBER=" & strRARuleNumber & _
                " ORDER BY APPLICATIONFACTFINDNUMBER DESC, STAGENUMBER DESC, RISKASSESSMENTSEQUENCENUMBER  DESC"
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = objIADOAssist.GetConnStr
    cmd.CommandText = strSQL
    Set rst = cmd.Execute(1, , adCmdText)

    Set objIClassDef = New omRAClassDef
    Set xmlClassDefDoc = objIClassDef.LoadPreviousRARuleOverrideData()

    Do While Not rst.EOF
        xmlResponse.appendChild objIDOAssist.GetXMLFromRecordSet(rst, xmlClassDefDoc)
        rst.MoveNext
    Loop

        rst.Close

    Set IRiskAssessmentDO_GetPreviousOverrideReasons = xmlOut.documentElement

    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objIADOAssist = Nothing
    Set objXmlAssist = Nothing
    Set objSQLAssist = Nothing
    Set cmd = Nothing
    Set rst = Nothing
    Set xmlDoc = Nothing
'DB END

' DB BMIDS00712 - Commented out.
'    Set xmlText = xmlOut.createElement("PREVIOUSOVERRIDEREASONS")
'    xmlText.Text = strOut
'    xmlResponse.appendChild xmlText
'
'    Set IRiskAssessmentDO_GetPreviousOverrideReasons = xmlResponse
' DB END
  
    'Tidy up and get out
        
    Exit Function

GetDataVbErr:
    
    ' Close the recordset if still open
        If Not rstRA Is Nothing Then
                If rstRA.State = adStateOpen Then
                        rstRA.Close
                End If
        End If
        Set rstRA = Nothing
        Set comSQL = Nothing

    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Function


Private Function IRiskAssessmentDO_GetRAApplicationStages(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
    On Error GoTo GetDataVbErr
    Const strFunctionName  As String = "GetRiskAssessment"

    Dim objIClassDef As IClassDefomRA
    Dim comSQL As ADODB.Command
    Dim objADOAssist As ADOAssist
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponse As IXMLDOMElement
    Dim xmlRAStages As IXMLDOMElement
    Dim xmlRAIn As IXMLDOMElement
    
    Dim strApplicationNumber As String
    Dim strAppFactFindNo As String
    Dim strStageNumber As String
    Dim rstStages As Recordset
    
    'Retrieve the minimum required input parameters
    'If vxmlTableElement.nodeName <> "RISKASSESSMENTAPPLICATIONSTAGELIST" Then
    '    Set xmlRAIn = m_objXmlAssist.GetMandatoryNode(vxmlTableElement, "RISKASSESSMENTAPPLICATIONSTAGELIST")
    'Else
        Set xmlRAIn = vxmlTableElement
    'End If
    
    strApplicationNumber = m_objXmlAssist.GetMandatoryElementText(xmlRAIn, "APPLICATIONNUMBER")
    strAppFactFindNo = m_objXmlAssist.GetElementText(xmlRAIn, "APPLICATIONFACTFINDNUMBER")
    
    'MAR1315 Remove stage number as a mandatory element - it is not needed here.
    strStageNumber = m_objXmlAssist.GetElementText(xmlRAIn, "STAGENUMBER")
    
    'Create an output document and assign it
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponse = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponse
    
    'Run the command
    Set comSQL = New ADODB.Command
    With comSQL
        .ActiveConnection = m_objADOAssist.GetConnStr
        .CommandText = _
        "SELECT APPLICATIONNUMBER, STAGENUMBER, STAGENAME, APPLICATIONFACTFINDNUMBER FROM STAGE," & _
        " ( SELECT DISTINCT APPLICATIONNUMBER, APPLICATIONFACTFINDNUMBER, STAGENUMBER FROM" & _
        " RISKASSESSMENT WHERE APPLICATIONNUMBER = '" & strApplicationNumber & "') R" & _
        " WHERE STAGE.STAGEID = R.STAGENUMBER" & _
        " ORDER BY STAGENUMBER DESC, APPLICATIONFACTFINDNUMBER DESC"
        Set rstStages = .Execute
    End With
    
    Set xmlRAStages = xmlOut.createElement("RISKASSESSMENTAPPLICATIONSTAGELIST")
    m_objIDOAssist.GetXMLFromWholeRecordset rstStages, StageClassDef, xmlRAStages
    Set IRiskAssessmentDO_GetRAApplicationStages = xmlRAStages
    rstStages.Close

    Set rstStages = Nothing
    Set xmlRAStages = Nothing
        Set comSQL = Nothing
       
    m_objXmlAssist.WriteXMLToFile (xmlOut.xml)
    
    'Tidy up and get out
    Exit Function

GetDataVbErr:
    
    ' Close the recordset if still open
        If Not rstStages Is Nothing Then
                If rstStages.State = adStateOpen Then
                        rstStages.Close
                End If
        End If
    Set rstStages = Nothing
        Set comSQL = Nothing

    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Private Function IRiskAssessmentDO_GetRiskAssessment(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
    
    On Error GoTo GetDataVbErr
    Const strFunctionName  As String = "GetRiskAssessment"

    Dim objIClassDef As IClassDefomRA
    Dim comSQL As ADODB.Command
    Dim objADOAssist As ADOAssist
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponse As IXMLDOMElement
    Dim xmlRA As IXMLDOMElement
    Dim xmlRARules As IXMLDOMElement
    Dim xmlRARule As IXMLDOMElement
    Dim xmlRAIn As IXMLDOMElement
    
    Dim strApplicationNumber As String
    Dim strAppFactFindNo As String
    Dim strStageNumber As String
    Dim strRASequenceNumber As String
    Dim rstRA As Recordset
    Dim rstRARules As Recordset
    Dim strRACount As String
    Dim strRAOffset As String
    Dim xmlNode As IXMLDOMNode
    
    'Retrieve the minimum required input parameters
    If vxmlTableElement.nodeName <> "RISKASSESSMENT" Then
        Set xmlRAIn = m_objXmlAssist.GetMandatoryNode(vxmlTableElement, "RISKASSESSMENT")
    Else
        Set xmlRAIn = vxmlTableElement
    End If
    
    strApplicationNumber = m_objXmlAssist.GetMandatoryElementText(xmlRAIn, "APPLICATIONNUMBER")
    strAppFactFindNo = m_objXmlAssist.GetElementText(xmlRAIn, "APPLICATIONFACTFINDNUMBER")
    strStageNumber = m_objXmlAssist.GetElementText(xmlRAIn, "STAGENUMBER")                        'MAR550
    strRASequenceNumber = m_objXmlAssist.GetElementText(xmlRAIn, "RISKASSESSMENTSEQUENCENUMBER")
    
    'Create an output document and assign it
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponse = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponse
    
    'Run the command
    Set comSQL = New ADODB.Command
    
     'MAR550
     'If StageNumber is missing then get latest otherwise get the stated one
    Dim strStageClause As String
    
    If strStageNumber <> vbNullString Then
        strStageClause = strStageNumber
    Else
        strStageClause = "(SELECT MAX(STAGENUMBER) FROM RISKASSESSMENT where APPLICATIONNUMBER='" & strApplicationNumber & "'" & _
            "AND APPLICATIONFACTFINDNUMBER=" & strAppFactFindNo & ")"
    End If
        
    'If RASeqeuencenumber is missing then get latest otherwise get the stated one
    'MAR550 Use the stage number derived above
    Dim strRAClause As String
    
    If strRASequenceNumber <> vbNullString Then
        strRAClause = strRASequenceNumber
    Else
        strRAClause = "(SELECT MAX(RISKASSESSMENTSEQUENCENUMBER) FROM RISKASSESSMENT where APPLICATIONNUMBER='" & strApplicationNumber & "'" & _
            "AND APPLICATIONFACTFINDNUMBER=" & strAppFactFindNo
            '& _

            If Len(strStageClause) > 0 Then
               strRAClause = strRAClause & " AND STAGENUMBER=" & strStageClause
            End If
            strRAClause = strRAClause & ")"
    End If
    
    ' Peter Edney - 05/04/2006
    ' MAR1563 / UAT 2194
    ' Get the total number of cases
    strRACount = "(Select Count(*)  From RISKASSESSMENT where APPLICATIONNUMBER='" & strApplicationNumber & _
                "' AND APPLICATIONFACTFINDNUMBER=" & strAppFactFindNo & " ) as RACOUNT," & vbCrLf
    ' Get the total number of case before the stage
    strRAOffset = "(Select Count(*)  From RISKASSESSMENT where APPLICATIONNUMBER='" & strApplicationNumber & _
                "' AND APPLICATIONFACTFINDNUMBER=" & strAppFactFindNo & _
                "AND STAGENUMBER<" & strStageClause & ") as RAOFFSET " & vbCrLf
                
    'MAR1291 Use stage
    With comSQL
        .ActiveConnection = m_objADOAssist.GetConnStr
        .CommandType = adCmdText
        .CommandText = "Select *, " & vbCrLf & _
                strRACount & strRAOffset & _
                " From RISKASSESSMENT where APPLICATIONNUMBER='" & strApplicationNumber & "'" & vbCrLf & _
                " AND APPLICATIONFACTFINDNUMBER=" & strAppFactFindNo & " AND " & vbCrLf & _
                "STAGENUMBER=" & strStageClause & " AND " & vbCrLf & _
                "RISKASSESSMENTSEQUENCENUMBER=" & strRAClause

        Set rstRA = .Execute
    End With
    
    ' A Risk assessment block is always returned but it may be empty
    If Not rstRA.BOF Then
        Set xmlRA = m_objIDOAssist.GetXMLFromRecordSet(rstRA, m_objIClassDef.LoadRiskAssessmentData, xmlResponse)
        
        'MAR550 Get the stage number from the returned XML
        If Not xmlRA.selectSingleNode("//RISKASSESSMENT") Is Nothing Then
            strStageNumber = xmlRA.selectSingleNode("//RISKASSESSMENT/STAGENUMBER").Text
            
            ' Peter Edney - 05/04/2006
            ' MAR1563 / UAT 2194
            Set xmlNode = xmlOut.createElement("RACOUNT")
            xmlNode.Text = rstRA.fields("RACOUNT").Value
            xmlRA.selectSingleNode("//RISKASSESSMENT").appendChild xmlNode
            Set xmlNode = xmlOut.createElement("RAOFFSET")
            xmlNode.Text = rstRA.fields("RAOFFSET").Value
            xmlRA.selectSingleNode("//RISKASSESSMENT").appendChild xmlNode
            
        End If
                
        'Now process the rules associated with this risk assessment - put the highest risks at the top.
        ' DM 16/05/01 SYS2314
        #If GENERIC_SQL Then
            With comSQL
                .ActiveConnection = m_objADOAssist.GetConnStr
                .CommandType = adCmdStoredProc
                .CommandText = "usp_GetRiskAssessment"
                .Prepared = True
            End With
           
            ' Need to set the precision for ADO NUMERIC types for SQL Server, this has no effect for Oracle
            Dim tmpParam As ADODB.Parameter
            comSQL.Parameters.Append comSQL.CreateParameter("p_ApplicationNumber", adBSTR, adParamInput, Len(strApplicationNumber), strApplicationNumber)
            ' Convert the string to a numeric type.
            Set tmpParam = comSQL.CreateParameter("p_ApplicationFactFindNumber", adNumeric, adParamInput, , CInt(strAppFactFindNo))
            tmpParam.Precision = 5
            comSQL.Parameters.Append tmpParam
            ' convert these
    
            Set tmpParam = comSQL.CreateParameter("p_StageNumber", adNumeric, adParamInput, , CInt(strStageNumber))
            tmpParam.Precision = 5
            comSQL.Parameters.Append tmpParam
            Set tmpParam = comSQL.CreateParameter("p_RiskAssessmentSequenceNumber", adNumeric, adParamInput, , CInt(rstRA.fields("RISKASSESSMENTSEQUENCENUMBER").Value))
            tmpParam.Precision = 5
            comSQL.Parameters.Append tmpParam
            
            Set rstRARules = comSQL.Execute

        ' Convert the above
        #Else
            comSQL.CommandText = "Select r.*," & _
                    "ro.USERID, ro.RAOVERRIDEDATETIME, ro.RAOVERRIDEREASON, ro.RAOVERRIDEAPPROVALREFERENCE " & _
                        " From RISKASSESSMENTRULE r, RISKASSESSMENTRULEOVERRIDE ro where " & _
                        "r.APPLICATIONNUMBER='" & strApplicationNumber & "'" & _
                        " AND r.APPLICATIONFACTFINDNUMBER=" & strAppFactFindNo & _
                        " AND r.STAGENUMBER=" & strStageNumber & " AND " & _
                        "r.RISKASSESSMENTSEQUENCENUMBER=" & rstRA.fields("RISKASSESSMENTSEQUENCENUMBER").Value & " AND " & _
                        "r.SCORECARDIND = 0 AND " & _
                        "ro.APPLICATIONNUMBER(+)=r.APPLICATIONNUMBER" & _
                        " AND ro.APPLICATIONFACTFINDNUMBER(+)=r.APPLICATIONFACTFINDNUMBER" & _
                        " AND ro.STAGENUMBER(+)=r.STAGENUMBER AND " & _
                        "ro.RISKASSESSMENTSEQUENCENUMBER(+)=r.RISKASSESSMENTSEQUENCENUMBER AND " & _
                        "ro.RARULENUMBER(+)=r.RARULENUMBER ORDER BY RARULESCORE DESC"
            Set rstRARules = comSQL.Execute
        #End If
        
        'Copy the Rules to the output
        Set xmlRARules = xmlOut.createElement("RISKASSESSMENTRULELIST")
        xmlRA.appendChild xmlRARules
    
        Do While Not rstRARules.EOF
            Set xmlRARule = m_objIDOAssist.GetXMLFromRecordSet(rstRARules, m_objIClassDef.LoadRARuleData, xmlRARules)
            
            xmlRARules.appendChild xmlRARule
            Dim XMLOverride As IXMLDOMElement
            Dim XMLOverrideText As IXMLDOMElement
            
            Set XMLOverride = m_objIDOAssist.GetXMLFromRecordSet(rstRARules, m_objIClassDef.LoadRARuleOverrideData, xmlRARules)
            
            Set XMLOverrideText = xmlRARule.ownerDocument.createElement("RAOVERRIDEREASONEX")
            XMLOverrideText.Text = FormatOverrideReason(rstRARules)
            
            xmlRARule.appendChild XMLOverride
            rstRARules.MoveNext
        Loop
                rstRARules.Close
        Set rstRARules = Nothing
    Else
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound, "No Risk Assessment found"
    End If
    
    Set IRiskAssessmentDO_GetRiskAssessment = xmlRA
       
    'Tidy up and get out
        
    Set comSQL = Nothing
        Exit Function

GetDataVbErr:
    
    ' Close the recordset if still open
        If Not rstRARules Is Nothing Then
                If rstRARules.State = adStateOpen Then
                        rstRARules.Close
                End If
        End If
        Set rstRARules = Nothing
    Set comSQL = Nothing

    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Sub IRiskAssessmentDO_CreateRuleOverride(ByVal vxmlTableElement As IXMLDOMElement)
    On Error GoTo GetDataVbErr
    Dim xmlDate As IXMLDOMElement
    Dim objSQLAssist As SQLAssist
    Set objSQLAssist = New SQLAssist
    
    Const strFunctionName  As String = "CreateRuleOverride"

    'Add the datea to the table
       
    Set xmlDate = vxmlTableElement.ownerDocument.createElement("RAOVERRIDEDATETIME")
    xmlDate.Text = objSQLAssist.DateTimeToString(Now())
    vxmlTableElement.appendChild xmlDate
    
    'Update the rule
    m_objIDOAssist.Create vxmlTableElement, m_objIClassDef.LoadRARuleOverrideData
    
    'Update the underlying loan
    IRiskAssessmentDO_UpdateRAUnderwriterDecision vxmlTableElement

    Exit Sub

GetDataVbErr:
    
    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
           
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Function IRiskAssessmentDO_GetRiskAssessmentRule(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
    On Error GoTo GetDataVbErr
    Const strFunctionName  As String = "GetRiskAssessment"

    Dim objIClassDef As IClassDefomRA
    Dim xmlRAIn As IXMLDOMElement
    
    'Retrieve the minimum required input parameters
    If vxmlTableElement.nodeName <> "RISKASSESSMENTRULE" Then
        Set xmlRAIn = m_objXmlAssist.GetMandatoryNode(vxmlTableElement, "RISKASSESSMENTRULE")
    Else
        Set xmlRAIn = vxmlTableElement
    End If
    
    Set objIClassDef = New omRAClassDef
    
    Set IRiskAssessmentDO_GetRiskAssessmentRule = m_objIDOAssist.GetData(xmlRAIn, objIClassDef.LoadRARuleData)
    
    Exit Function

GetDataVbErr:
    
    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
           
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Private Function IRiskAssessmentDO_GetUserRiskAssessmentAuthority(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
On Error GoTo GetDataVbErr
    Const strFunctionName  As String = "GetRiskAssessment"

    Dim objIClassDef As IClassDefomRA
    Dim objSQLAssist As SQLAssist
    Dim comSQL As ADODB.Command
    Dim objADOAssist As ADOAssist
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlRALevel As IXMLDOMElement
    Dim rstRALevel As Recordset
    Dim strUserId As String
    Dim xmlUSERIDattrib As IXMLDOMAttribute
    
    'Use the standard userid if supplied.
    If vxmlTableElement.nodeName = "REQUEST" Then
        Set xmlUSERIDattrib = vxmlTableElement.Attributes.getNamedItem("USERID")
        If Not xmlUSERIDattrib Is Nothing Then
            strUserId = xmlUSERIDattrib.Value
        End If
    End If
    
    'If there wasn't a standard userid then look for it explicitly
    If strUserId = vbNullString Then
        If vxmlTableElement.nodeName = "USERID" Then
            strUserId = vxmlTableElement.Text
        Else
            strUserId = m_objXmlAssist.GetMandatoryElementText(vxmlTableElement, "USERID")
        End If
    End If
    
    Set objSQLAssist = New SQLAssist

    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    
    'Run the command. Look for the highest active RiskAssessment level for a user
    Set comSQL = New ADODB.Command
    With comSQL
        .ActiveConnection = m_objADOAssist.GetConnStr
        .CommandType = adCmdText
        .CommandText = _
            "SELECT MAX(COMPETENCY.RISKASSESSMENTMANDATE) AS RISKASSESSMENTMANDATE FROM USERCOMPETENCY,COMPETENCY" & _
            " WHERE USERCOMPETENCY.USERID='" & strUserId & _
            "' AND COMPETENCY.COMPETENCYACTIVEFROM <=" & objSQLAssist.FormatDateString(Format(Now(), "dd/mm/yyyy hh:mm:ss")) & _
            " AND COMPETENCY.COMPETENCYTYPE=USERCOMPETENCY.COMPETENCYTYPE"

        Set rstRALevel = .Execute
    End With
    
    'Copy the basic stuff to the output
    Set xmlRALevel = xmlOut.createElement("RISKASSESSMENTMANDATE")
    
    If Not IsNull(rstRALevel.fields("RISKASSESSMENTMANDATE").Value) Then
        xmlRALevel.Text = rstRALevel.fields("RISKASSESSMENTMANDATE").Value
    Else
        xmlRALevel.Text = "0"
    End If
    rstRALevel.Close
        Set rstRALevel = Nothing

    Set IRiskAssessmentDO_GetUserRiskAssessmentAuthority = xmlRALevel
    
    Exit Function

GetDataVbErr:
    
    ' Close the recordset if still open
        If Not rstRALevel Is Nothing Then
                If rstRALevel.State = adStateOpen Then
                        rstRALevel.Close
                End If
        End If
        Set rstRALevel = Nothing

    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
            
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Function


Private Function FormatOverrideReason(ByVal vrstData As Recordset, Optional ByVal vblnWithRef As Boolean = False) As String
    Dim strRef
    Const strHDR = "------------------"
    'Add in a reference description if requested
    If vblnWithRef And Not IsNull(vrstData.fields("RAOVERRIDEAPPROVALREFERENCE").Value) Then
        strRef = "Approval Reference: " & vrstData.fields("RAOVERRIDEAPPROVALREFERENCE").Value & vbCrLf
    End If
    
    FormatOverrideReason = strHDR & " " & _
        Format$(vrstData.fields("RAOVERRIDEDATETIME").Value, "dd/mm/yyyy") & " " & _
        vrstData.fields("USERID").Value & " " & _
        strHDR & vbCrLf & strRef & _
        vrstData.fields("RAOVERRIDEREASON").Value

End Function


Private Sub IRiskAssessmentDO_UpdateRAUnderwriterDecision(ByVal vxmlTableElement As IXMLDOMElement)
    On Error GoTo GetDataVbErr
    Dim xmlDate As IXMLDOMElement
    Dim xmlRequest As FreeThreadedDOMDocument40
    Dim xmlTable As IXMLDOMElement
    
    Dim strAppNo As String
    Dim strAppFactFindNo As String
    Dim strStageNo As String
    Dim strRASeqNo As String
    Dim strUWDecision As String
    
    Const strFunctionName  As String = "UpdateRAUnderwriterDecision"

    strAppNo = m_objXmlAssist.GetMandatoryElementText(vxmlTableElement, "APPLICATIONNUMBER")
    strAppFactFindNo = m_objXmlAssist.GetMandatoryElementText(vxmlTableElement, "APPLICATIONFACTFINDNUMBER")
    strStageNo = m_objXmlAssist.GetMandatoryElementText(vxmlTableElement, "STAGENUMBER")
    strRASeqNo = m_objXmlAssist.GetMandatoryElementText(vxmlTableElement, "RISKASSESSMENTSEQUENCENUMBER")
    strUWDecision = m_objXmlAssist.GetElementText(vxmlTableElement, "UNDERWRITERDECISION")
    
    'If the UW Decision is missing it will need to be calculated from the global parameters
    If strUWDecision = vbNullString Then
        'Now recalculate the risk on the loan. This is performed by looking at all scores which are
        'not overriden will be counted
        Dim comSQL As Command
        ' DM 16/05/01 SYS2314
        #If GENERIC_SQL Then
            Dim adoParameterScore As ADODB.Parameter
        #Else
            Dim rsScore As Recordset
        #End If
        Dim strScore As String
        Dim xmlResult As FreeThreadedDOMDocument40
        
        Set comSQL = New Command
        comSQL.ActiveConnection = m_objADOAssist.GetConnStr
        
        ' DM 16/05/01 SYS2314
        #If GENERIC_SQL Then
            ' implement the stored proc call in here
            With comSQL
                .CommandType = adCmdStoredProc
                .CommandText = "usp_GetRARuleScore"
                .Prepared = True
            End With
            
            ' Set up the input parameters
            Dim tmpParam As ADODB.Parameter
            comSQL.Parameters.Append comSQL.CreateParameter("p_strAppNo", adBSTR, adParamInput, Len(strAppNo), strAppNo)
            Set tmpParam = comSQL.CreateParameter("p_iAppFactFindNo", adNumeric, adParamInput, , CInt(strAppFactFindNo))
            tmpParam.Precision = 5
            comSQL.Parameters.Append tmpParam
            Set tmpParam = comSQL.CreateParameter("p_iStageNo", adNumeric, adParamInput, , CInt(strStageNo))
            tmpParam.Precision = 5
            comSQL.Parameters.Append tmpParam
            Set tmpParam = comSQL.CreateParameter("p_iRASeqNo", adNumeric, adParamInput, , CInt(strRASeqNo))
            tmpParam.Precision = 5
            comSQL.Parameters.Append tmpParam
            ' DM SYS2511 Changed datatype adVarChar
            Set adoParameterScore = comSQL.CreateParameter("p_strRARuleScore", adVarChar, adParamOutput, 10)
            comSQL.Parameters.Append adoParameterScore
            
            comSQL.Execute
            
            strScore = adoParameterScore.Value
            Set adoParameterScore = Nothing

        #Else
                
            comSQL.CommandType = adCmdText
            comSQL.CommandText = _
            "select SUM(RARULESCORE) AS SCORE FROM (" & _
            "select r.*, o.USERID FROM RISKASSESSMENTRULE r, RISKASSESSMENTRULEOVERRIDE o " & _
            "where r.APPLICATIONNUMBER='" & strAppNo & "' and r.APPLICATIONFACTFINDNUMBER=" & strAppFactFindNo & " and " & _
            "r.STAGENUMBER=" & strStageNo & " And r.RISKASSESSMENTSEQUENCENUMBER =" & strRASeqNo & _
            " and r.RARULESCORE>0 and " & _
            "o.APPLICATIONNUMBER (+)=r.APPLICATIONNUMBER and o.APPLICATIONFACTFINDNUMBER (+)=r.APPLICATIONFACTFINDNUMBER and " & _
            "o.STAGENUMBER (+)=r.STAGENUMBER and o.RISKASSESSMENTSEQUENCENUMBER (+)=r.RISKASSESSMENTSEQUENCENUMBER " & _
            "and o.RARULENUMBER(+)=r.RARULENUMBER) where USERID is null"
            
            Set rsScore = comSQL.Execute
        
        #End If
        
        Dim objGP As GlobalBandedParameterDO
        If Not m_objContext Is Nothing Then
            Set objGP = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".GlobalBandedParameterDO")
        Else
            Set objGP = New GlobalBandedParameterDO
        End If
        
        ' DM 16/05/01 SYS2314
        #If GENERIC_SQL Then
        #Else
        If Not IsNull(rsScore!SCORE) Then strScore = rsScore!SCORE Else strScore = "0"
        rsScore.Close
                Set rsScore = Nothing
        #End If
        
        Set xmlResult = _
            m_objXmlAssist.load(objGP.GetCurrentParameter("RiskAssessmentDecision", strScore), TypeName(Me), TypeName(Me))
        
        strUWDecision = m_objXmlAssist.GetMandatoryElementText(xmlResult.firstChild, "AMOUNT")
        
    End If
    
    
    Set xmlRequest = New FreeThreadedDOMDocument40
    xmlRequest.validateOnParse = False
    xmlRequest.setProperty "NewParser", True
    xmlRequest.async = False
    xmlRequest.loadXML _
        "<RISKASSESSMENT>" & _
            "<APPLICATIONNUMBER>" & strAppNo & "</APPLICATIONNUMBER>" & _
            "<APPLICATIONFACTFINDNUMBER>" & strAppFactFindNo & "</APPLICATIONFACTFINDNUMBER>" & _
            "<STAGENUMBER>" & strStageNo & "</STAGENUMBER>" & _
            "<RISKASSESSMENTSEQUENCENUMBER>" & strRASeqNo & "</RISKASSESSMENTSEQUENCENUMBER>" & _
        "<UNDERWRITERDECISION>" & strUWDecision & "</UNDERWRITERDECISION>" & _
    "</RISKASSESSMENT>"
    
    m_objIDOAssist.Update xmlRequest.firstChild, m_objIClassDef.LoadRiskAssessmentData
    
        Set comSQL = Nothing
    Exit Sub

GetDataVbErr:
    
    ' Close the recordset if still open
        ' DM 16/05/01 SYS2314
        #If GENERIC_SQL Then
            Set adoParameterScore = Nothing
        #Else
        If Not rsScore Is Nothing Then
                If rsScore.State = adStateOpen Then
                        rsScore.Close
                End If
        End If
        Set rsScore = Nothing
        #End If

        Set comSQL = Nothing

    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
           
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

Private Function StageClassDef() As FreeThreadedDOMDocument40

    Const strFunctionName As String = "StageClassDef"

    Dim strXML As String
    
    strXML = strXML & "<TABLENAME>" & _
        "APPLICATIONSTAGE" & _
        "<PRIMARYKEY>APPLICATIONNUMBER<TYPE>dbdtString</TYPE></PRIMARYKEY>" & _
        "<PRIMARYKEY>APPLICATIONFACTFINDNUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>" & _
        "<PRIMARYKEY>STAGENUMBER<TYPE>dbdtInt</TYPE></PRIMARYKEY>" & _
        "<OTHERS>STAGENAME<TYPE>dbdtString</TYPE></OTHERS>" & _
     "</TABLENAME>"

    Set StageClassDef = m_objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
    

End Function

Private Function IRiskAssessmentDO_GetRiskAssessmentDecisionParameter( _
    ByVal vintFailCount As Integer) _
    As String
    
On Error GoTo GetRiskAssessmentDecisionParameterVbErr
    
    Dim rst As ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "SELECT AMOUNT FROM GLOBALBANDEDPARAMETER GBPARAM1 WHERE"
    strSQL = strSQL & " NAME = 'RiskAssessmentDecision'"
    strSQL = strSQL & " AND GBPARAM1.GBPARAMSTARTDATE ="
    strSQL = strSQL & " (SELECT MAX(GBPARAMSTARTDATE) FROM GLOBALBANDEDPARAMETER WHERE"
    strSQL = strSQL & " NAME = 'RiskAssessmentDecision'"
    strSQL = strSQL & " AND GBPARAMSTARTDATE <= " & m_objSQLAssist.GetSystemDate & ")"
    strSQL = strSQL & " AND GBPARAM1.HIGHBAND = "
    strSQL = strSQL & " (SELECT MIN(HIGHBAND) FROM GLOBALBANDEDPARAMETER WHERE"
    strSQL = strSQL & " NAME = 'RiskAssessmentDecision'"
    strSQL = strSQL & " AND HIGHBAND >= " & vintFailCount
    strSQL = strSQL & " AND GBPARAMSTARTDATE = GBPARAM1.GBPARAMSTARTDATE)"

    Set rst = adoGetRecordSet(strSQL)

    If Not rst Is Nothing Then
        If Not IsNull(rst.fields.Item(0).Value) Then
            IRiskAssessmentDO_GetRiskAssessmentDecisionParameter = _
                rst.fields.Item(0).Value
        End If
        
        rst.Close
        Set rst = Nothing
        
    End If
    
    Exit Function
    
GetRiskAssessmentDecisionParameterVbErr:
    
    ' Close the recordset if still open
        If Not rst Is Nothing Then
                If rst.State = adStateOpen Then
                        rst.Close
                End If
        End If
        Set rst = Nothing

    Err.Raise Err.Number, _
        "GetRiskAssessmentDecisionParameter." & Err.Source, _
        Err.Description, _
        Err.HelpFile, _
        Err.HelpContext

End Function

Private Function IRiskAssessmentDO_StoreRiskAssResults( _
    ByVal vobjXmlRequest As IXMLDOMNode, _
    ByVal vobjXmlRulesResponse As IXMLDOMNode) As IXMLDOMNode
    
On Error GoTo StoreRiskAssResultsVbErr
    
    Dim strApplicationNum As String, _
        strApplicationFactFindNumber As String, _
        strStageNumber As String, _
        strUserId As String, _
        strCaseSequenceNo As String
    Dim intRiskAssessmentSeq As Integer
    Dim domDoc As FreeThreadedDOMDocument40
    
    strUserId = _
        m_objXmlAssist.GetMandatoryAttribute(vobjXmlRequest, "USERID")
    
    strApplicationNum = _
        m_objXmlAssist.GetMandatoryElementText(vobjXmlRequest, ".//APPLICATIONNUMBER")
    
    strApplicationFactFindNumber = _
        m_objXmlAssist.GetMandatoryElementText(vobjXmlRequest, ".//APPLICATIONFACTFINDNUMBER")
    
    If (Not m_objXmlAssist.GetElementTextBln(vobjXmlRequest, ".//STAGENUMBER", strStageNumber)) Or _
        (Not m_objXmlAssist.GetElementTextBln(vobjXmlRequest, ".//CASESEQUENCENO", strCaseSequenceNo)) Then
        
        'SYS2074
        GetApplicationStageInformationForRA strApplicationNum, strApplicationFactFindNumber, strStageNumber, strCaseSequenceNo
    
    End If
    
    intRiskAssessmentSeq = _
        GetRiskAssessmentSeq( _
            strApplicationNum, _
            strApplicationFactFindNumber, _
            strStageNumber)
    
    ' SYS2074
    CreateRiskAssessment _
        strApplicationNum, _
        strApplicationFactFindNumber, _
        strStageNumber, _
        strCaseSequenceNo, _
        strUserId, _
        intRiskAssessmentSeq, _
        vobjXmlRulesResponse
    
    CreateRiskAssessmentRules _
        strApplicationNum, _
        strApplicationFactFindNumber, _
        strStageNumber, _
        intRiskAssessmentSeq, _
        vobjXmlRulesResponse

    Set domDoc = New FreeThreadedDOMDocument40
    domDoc.validateOnParse = False
    domDoc.setProperty "NewParser", True
    domDoc.async = False
    
    ' Results are stored under a new sequence number within the RISKASSESSMENT table
    ' This method returns all the key fields that describe the row that was created
    ' This information is then used to update the underwriter decision
    domDoc.loadXML "<RESPONSE TYPE=""SUCCESS"">" & _
        "<RISKASSESSMENT>" & _
            "<APPLICATIONNUMBER>" & strApplicationNum & "</APPLICATIONNUMBER>" & _
            "<APPLICATIONFACTFINDNUMBER>" & strApplicationFactFindNumber & "</APPLICATIONFACTFINDNUMBER>" & _
            "<STAGENUMBER>" & strStageNumber & "</STAGENUMBER>" & _
            "<RISKASSESSMENTSEQUENCENUMBER>" & CStr(intRiskAssessmentSeq) & "</RISKASSESSMENTSEQUENCENUMBER>" & _
        "</RISKASSESSMENT>" & _
        "</RESPONSE>"
        
    Set IRiskAssessmentDO_StoreRiskAssResults = domDoc.firstChild
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Exit Function
    
StoreRiskAssResultsVbErr:
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Err.Raise Err.Number, _
        "[RiskAssessmentDO].StoreRiskAssResults." & Err.Source, _
        Err.Description, _
        Err.HelpFile, _
        Err.HelpContext

End Function

Private Sub GetApplicationStageInformationForRA( _
    ByVal vstrApplicationNum As String, _
    ByVal vstrApplicationFactFindNum As String, _
    ByRef rstrStageNumber As String, _
    ByRef rstrCaseSequenceNo As String)
    
On Error GoTo GetApplicationStageVbErr
    
    Dim rst As ADODB.Recordset
    Dim strSQL As String
    
    'SYS2074
    strSQL = "select CASESEQUENCENO, Max(STAGENUMBER) From APPLICATIONSTAGE "
    strSQL = strSQL & "where APPLICATIONNUMBER = " & m_objSQLAssist.FormatString(vstrApplicationNum) & " And CASESEQUENCENO = "
    strSQL = strSQL & "(select max(CASESEQUENCENO) from APPLICATIONSTAGE where APPLICATIONNUMBER = " & m_objSQLAssist.FormatString(vstrApplicationNum)
    strSQL = strSQL & "and APPLICATIONFACTFINDNUMBER= " & vstrApplicationFactFindNum & ") group by CASESEQUENCENO "
    
'    strSQL = "select max(STAGENUMBER) from APPLICATIONSTAGE"
'    strSQL = strSQL & " where applicationnumber=" & m_objSQLAssist.FormatString(vstrApplicationNum)
'    strSQL = strSQL & " and applicationfactfindnumber =" & vstrApplicationFactFindNum

    Set rst = adoGetRecordSet(strSQL)

    If Not rst Is Nothing Then
        
        If Not IsNull(rst.fields.Item(0).Value) Then
             rstrCaseSequenceNo = CStr(rst.fields.Item(0).Value)
        End If
        
        If Not IsNull(rst.fields.Item(0).Value) Then
             rstrStageNumber = CStr(rst.fields.Item(1).Value)
        End If
        
        rst.Close
        Set rst = Nothing
        
    End If
    
    Exit Sub
    
GetApplicationStageVbErr:
    
    ' Close the recordset if still open
        If Not rst Is Nothing Then
                If rst.State = adStateOpen Then
                        rst.Close
                End If
        End If
        Set rst = Nothing

    Err.Raise Err.Number, _
        "GetRiskAssessmentSeq." & Err.Source, _
        Err.Description, _
        Err.HelpFile, _
        Err.HelpContext
    

End Sub


Private Function GetRiskAssessmentSeq( _
    ByVal vstrApplicationNum As String, _
    ByVal vstrApplicationFactFindNum As String, _
    ByVal vstrStageNumber As String _
    ) As Integer
    
On Error GoTo StoreRiskAssResultsVbErr
    
    Dim rst As ADODB.Recordset
    Dim strSQL As String
    
    GetRiskAssessmentSeq = 1
    
    ' DM    20/11/2002  BMIDS01022
    strSQL = "select max(RISKASSESSMENTSEQUENCENUMBER) from riskassessment WITH (NOLOCK)"
    strSQL = strSQL & " where applicationnumber=" & m_objSQLAssist.FormatString(vstrApplicationNum)
    strSQL = strSQL & " and applicationfactfindnumber =" & vstrApplicationFactFindNum
    strSQL = strSQL & " and stagenumber =" & vstrStageNumber

    Set rst = adoGetRecordSet(strSQL)

    If Not rst Is Nothing Then
        If Not IsNull(rst.fields.Item(0).Value) Then
            GetRiskAssessmentSeq = _
                GetRiskAssessmentSeq + rst.fields.Item(0).Value
        End If
        
        rst.Close
        Set rst = Nothing
        
    End If
    
    Exit Function
    
StoreRiskAssResultsVbErr:
    
    ' Close the recordset if still open
        If Not rst Is Nothing Then
                If rst.State = adStateOpen Then
                        rst.Close
                End If
        End If
        Set rst = Nothing

    Err.Raise Err.Number, _
        "GetRiskAssessmentSeq." & Err.Source, _
        Err.Description, _
        Err.HelpFile, _
        Err.HelpContext
    
End Function

Private Sub CreateRiskAssessment( _
    ByVal vstrApplicationNum As String, _
    ByVal vstrApplicationFactFindNum As String, _
    ByVal vstrStageNumber As String, _
    ByVal vstrCaseSequenceNo As String, _
    ByVal vstrUserId As String, _
    ByVal vstrSeqNo As String, _
    vobjXmlRulesResponse As IXMLDOMNode)

On Error GoTo CreateRiskAssessmentVbErr

    Dim strAttrib As String, _
        strSQL As String, _
        strNames As String, _
        strValues As String
    
    Dim intScore As Integer, _
        intDecision As Integer
    
    If xmlGetAttribute(vobjXmlRulesResponse, "FAILCOUNT", strAttrib) Then
        intScore = CInt(strAttrib)
    End If
    
    If xmlGetAttribute(vobjXmlRulesResponse, "DECISION", strAttrib) Then
        intDecision = CInt(strAttrib)
    End If
    
    strNames = "APPLICATIONNUMBER"
    strValues = m_objSQLAssist.FormatString(vstrApplicationNum)
    
    strNames = strNames & ", APPLICATIONFACTFINDNUMBER"
    strValues = strValues & ", " & vstrApplicationFactFindNum
    
    strNames = strNames & ", STAGENUMBER"
    strValues = strValues & ", " & vstrStageNumber
    
    ' SYS2074
    strNames = strNames & ", CASESEQUENCENO"
    strValues = strValues & ", " & vstrCaseSequenceNo
    
    strNames = strNames & ", RISKASSESSMENTSEQUENCENUMBER"
    strValues = strValues & ", " & vstrSeqNo
    
    strNames = strNames & ", USERID"
    strValues = strValues & ", " & m_objSQLAssist.FormatString(vstrUserId)
    
    strNames = strNames & ", RISKASSESSMENTDATETIME"
    strValues = strValues & ", " & m_objSQLAssist.FormatDate(Now())
    
    strNames = strNames & ", RISKASSESSMENTSCORE"
    strValues = strValues & ", " & intScore
    
    strNames = strNames & ", RISKASSESSMENTDECISION"
    strValues = strValues & ", " & intDecision
    
    strSQL = _
        "INSERT INTO RISKASSESSMENT" & _
        " (" & strNames & ")" & _
        " VALUES (" & strValues & ")"
        
    m_objADOAssist.ExecuteSQLCommand strSQL, False
        
    Exit Sub
    
CreateRiskAssessmentVbErr:
    
    Err.Raise Err.Number, _
        "CreateRiskAssessment." & Err.Source, _
        Err.Description, _
        Err.HelpFile, _
        Err.HelpContext

End Sub

Private Sub CreateRiskAssessmentRules( _
    ByVal vstrApplicationNum As String, _
    ByVal vstrApplicationFactFindNum As String, _
    ByVal vstrStageNumber As String, _
    ByVal vstrSeqNo As String, _
    vobjXmlRulesResponse As IXMLDOMNode)

On Error GoTo CreateRiskAssessmentRulesVbErr

    Dim objXmlNodes As IXMLDOMNodeList

    Dim strAttrib As String
    Dim intIndex As Integer
        
    Set objXmlNodes = vobjXmlRulesResponse.selectNodes("RULE")
    
        For intIndex = 0 To (objXmlNodes.length - 1)
        
            If xmlGetAttribute(objXmlNodes.Item(intIndex), "NUMBER", strAttrib) Then
            
                CreateRiskAssessmentRule _
                    vstrApplicationNum, _
                    vstrApplicationFactFindNum, _
                    vstrStageNumber, _
                    vstrSeqNo, _
                    strAttrib, _
                    objXmlNodes.Item(intIndex)
            
            End If
        
        Next

    Exit Sub

CreateRiskAssessmentRulesVbErr:
    
    Err.Raise Err.Number, _
        "CreateRiskAssessmentRules." & Err.Source, _
        Err.Description, _
        Err.HelpFile, _
        Err.HelpContext

End Sub


Private Sub CreateRiskAssessmentRule( _
    ByVal vstrApplicationNum As String, _
    ByVal vstrApplicationFactFindNum As String, _
    ByVal vstrStageNumber As String, _
    ByVal vstrSeqNo As String, _
    ByVal vstrRuleNumber As String, _
    vobjXmlRulesNode As IXMLDOMNode)

    Dim strAttrib As String, _
        strSQL As String, _
        strNames As String, _
        strValues As String, _
        strRuleName As String
    Dim intScore As Integer, _
        intResult As Integer
    Dim strScoreCardInd As String
    
    strNames = "APPLICATIONNUMBER"
    strValues = m_objSQLAssist.FormatString(vstrApplicationNum)
    
    strNames = strNames & ", APPLICATIONFACTFINDNUMBER"
    strValues = strValues & ", " & vstrApplicationFactFindNum
    
    strNames = strNames & ", STAGENUMBER"
    strValues = strValues & ", " & vstrStageNumber
    
    strNames = strNames & ", RISKASSESSMENTSEQUENCENUMBER"
    strValues = strValues & ", " & vstrSeqNo
    
    strNames = strNames & ", RARULENUMBER"
    strValues = strValues & ", " & vstrRuleNumber
    
    xmlGetAttribute vobjXmlRulesNode, "TEXTONE", strRuleName
    strNames = strNames & ", RARULENAME"
    strValues = strValues & ", " & m_objSQLAssist.FormatString(strRuleName)
    
    If xmlGetAttribute(vobjXmlRulesNode, "SCORE", strAttrib) Then
        intScore = CInt(strAttrib)
    End If
    
    strNames = strNames & ", RARULESCORE"
    strValues = strValues & ", " & intScore
    
    If intScore > 0 Then
        intResult = 2
    Else
        intResult = 1
    End If
    
    strNames = strNames & ", RARULERESULT"
    strValues = strValues & ", " & intResult
    
    If xmlGetAttribute(vobjXmlRulesNode, "TEXTTWO", strAttrib) Then
        strNames = strNames & ", RARULEVALUE"
        strValues = strValues & ", " & m_objSQLAssist.FormatString(strAttrib)
    End If
    
    'MV - 19/07/2002 - BMIDS00219 - Amend ScoreCardInd
    If xmlGetAttribute(vobjXmlRulesNode, "SCORECARDIND", strScoreCardInd) Then
        strNames = strNames & ", SCORECARDIND"
        strValues = strValues & ", " & strScoreCardInd
    End If
    
    strSQL = _
        "INSERT INTO RISKASSESSMENTRULE" & _
        " (" & strNames & ")" & _
        " VALUES (" & strValues & ")"
        
    m_objADOAssist.ExecuteSQLCommand strSQL, False

End Sub

Private Function adoGetRecordSet(ByVal vstrSQL As String) As ADODB.Recordset
On Error GoTo GetRecordSetVbErr

    Dim conn As ADODB.Connection
    Dim rst As ADODB.Recordset
    
    Set conn = New ADODB.Connection
    
    conn.ConnectionString = m_objADOAssist.GetConnStr
    conn.open
    
    Set rst = New ADODB.Recordset
    
    rst.CursorLocation = adUseClient
    rst.open vstrSQL, conn, adOpenStatic, adLockReadOnly, adCmdText

    ' disconnect RecordSet
    Set rst.ActiveConnection = Nothing
        conn.Close

    If Not rst.EOF Then
        rst.MoveFirst
        Set adoGetRecordSet = rst
    End If
    
    Set rst = Nothing
    Set conn = Nothing
    
    Exit Function
    
GetRecordSetVbErr:
    
    ' Close the connection if still open
        If Not conn Is Nothing Then
                If conn.State = adStateOpen Then
                        conn.Close
                End If
        End If

    Set rst = Nothing
    Set conn = Nothing
    
    Err.Raise _
        Err.Number, _
        "[RiskAssessmentDO].GetRecordSet." & Err.Source, _
        Err.Description

End Function

Private Function xmlGetAttribute( _
    ByVal vobjNode As IXMLDOMNode, _
    ByVal vstrAttribName As String, _
    ByRef rstrAttribText As String) _
    As Boolean
    
    xmlGetAttribute = False
    
    Dim strValue As String
    
    If Not vobjNode.Attributes.getNamedItem(vstrAttribName) Is Nothing Then
        If Not vobjNode.Attributes.getNamedItem(vstrAttribName).Text = Empty Then
            rstrAttribText = vobjNode.Attributes.getNamedItem(vstrAttribName).Text
            xmlGetAttribute = True
        End If
    End If

End Function

Private Function IRiskAssessmentDO_GetLatestRiskAssessmentScore(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for all instances of the persistant data associated with
'   this data object for the values supplied
' pass:
'   vxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo GetLatestRiskAssessmentScoreVbErr

    Const strFunctionName As String = "IRiskAssessmentDO_GetLatestRiskAssessmentScore"
    
        
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim objIADOAssist As IADOAssist
    Dim objXmlAssist As XMLAssist
    Dim xmlDataNode As IXMLDOMElement
    Dim xmlElement As IXMLDOMElement
    
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
    
    Set objXmlAssist = New XMLAssist
    
    ' Check the correct keys have been passed in
    strApplicationNumber = objXmlAssist.GetNodeValue(vxmlTableElement, _
                                                     ".//APPLICATIONNUMBER")
    strApplicationFFNumber = objXmlAssist.GetNodeValue(vxmlTableElement, _
                                                       ".//APPLICATIONFACTFINDNUMBER")
    
    Set objIADOAssist = New ADOAssist
    
    Set adoConnection = New ADODB.Connection
    Set adoCommand = New ADODB.Command
    Set adoRecordset = New ADODB.Recordset
        
    Dim strSQL As String
    
    strSQL = "SELECT RA.RISKASSESSMENTSCORE FROM RISKASSESSMENT RA, " & _
             "(SELECT APPLICATIONNUMBER, APPLICATIONFACTFINDNUMBER, MAX(STAGENUMBER) " & _
             "STAGENUMBER FROM RISKASSESSMENT GROUP BY APPLICATIONNUMBER, " & _
             "APPLICATIONFACTFINDNUMBER) RASTAGE, " & _
             "(SELECT APPLICATIONNUMBER, APPLICATIONFACTFINDNUMBER, STAGENUMBER, " & _
             "MAX(RISKASSESSMENTSEQUENCENUMBER) RISKASSESSMENTSEQUENCENUMBER " & _
             "FROM RISKASSESSMENT GROUP BY APPLICATIONNUMBER, APPLICATIONFACTFINDNUMBER, " & _
             "STAGENUMBER) RASEQUENCE WHERE RASTAGE.APPLICATIONNUMBER = RASEQUENCE.APPLICATIONNUMBER " & _
             "AND RASTAGE.APPLICATIONFACTFINDNUMBER = RASEQUENCE.APPLICATIONFACTFINDNUMBER " & _
             "AND RASTAGE.STAGENUMBER=RASEQUENCE.STAGENUMBER " & _
             "AND RA.APPLICATIONNUMBER = RASEQUENCE.APPLICATIONNUMBER " & _
             "AND RA.APPLICATIONFACTFINDNUMBER = RASEQUENCE.APPLICATIONFACTFINDNUMBER " & _
             "AND RA.STAGENUMBER=RASEQUENCE.STAGENUMBER " & _
             "AND RA.RISKASSESSMENTSEQUENCENUMBER = RASEQUENCE.RISKASSESSMENTSEQUENCENUMBER " & _
             "AND RA.APPLICATIONNUMBER = ? AND RA.APPLICATIONFACTFINDNUMBER = ?"
    
    adoCommand.CommandText = strSQL
    
    ' Set up the parameters to be substituted into the SQL
    Set adoParameter = adoCommand.CreateParameter("ApplicationNumber", adBSTR, adParamInput, Len(strApplicationNumber), strApplicationNumber)
    adoCommand.Parameters.Append adoParameter

    Set adoParameter = adoCommand.CreateParameter("ApplicationFactFindNumber", adInteger, adParamInput)
    adoCommand.Parameters.Append adoParameter
    adoCommand(1) = strApplicationFFNumber
    
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
         
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True

    Set xmlDataNode = xmlOut.createElement("RISKASSESSMENT")
    xmlOut.appendChild xmlDataNode

    ' Convert recordset to XML
    Set xmlElement = xmlOut.createElement("RISKASSESSMENTSCORE")
    xmlElement.Text = adoRecordset.fields(0).Value
    xmlDataNode.appendChild xmlElement
    
    adoRecordset.Close
    
    Set IRiskAssessmentDO_GetLatestRiskAssessmentScore = xmlDataNode
    
    Set xmlOut = Nothing
    Set objIADOAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlDataNode = Nothing
    Set adoConnection = Nothing
    Set adoCommand = Nothing
    Set adoRecordset = Nothing
    Set adoParameter = Nothing
    Set xmlElement = Nothing
    
    m_objContext.SetComplete
    
    
    Exit Function

GetLatestRiskAssessmentScoreVbErr:
    
    
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
    
    Set xmlOut = Nothing
    Set objIADOAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlDataNode = Nothing
    Set adoConnection = Nothing
    Set adoCommand = Nothing
    Set adoRecordset = Nothing
    Set adoParameter = Nothing
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

' SR 09/08/2004 : BMIDS815 - new method
Private Function IRiskAssessmentDO_CheckRiskAssessmentExistsForApplication( _
                            ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode

On Error GoTo CheckRiskAssessmentExistsForApplicationVbErr

    Const strFunctionName As String = "IRiskAssessmentDO_CheckRiskAssessmentExistsForApplication"
   
    Dim strApplicationNumber As String, strWhereCondition As String
    Dim blnRAExists As Boolean
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponse As IXMLDOMElement, xmlElem As IXMLDOMElement
    
    'Create an output document and assign it
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Set xmlResponse = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponse
        
    strApplicationNumber = m_objXmlAssist.GetMandatoryElementText(vxmlTableElement, "APPLICATIONNUMBER")
    strWhereCondition = " APPLICATIONNUMBER = " & m_objSQLAssist.FormatString(strApplicationNumber)
    blnRAExists = m_objADOAssist.CheckRecordExists("RISKASSESSMENT", strWhereCondition)
    
    Set xmlElem = xmlOut.createElement("RISKASSESSMENTEXISTS")
    xmlResponse.appendChild xmlElem
    If blnRAExists Then
        xmlElem.Text = 1
    Else
        xmlElem.Text = 0
    End If
    
    Set IRiskAssessmentDO_CheckRiskAssessmentExistsForApplication = xmlResponse
    
    Set xmlOut = Nothing
    Set xmlElem = Nothing
    Set xmlResponse = Nothing
    
    Exit Function
CheckRiskAssessmentExistsForApplicationVbErr:

    Set xmlOut = Nothing
    Set xmlResponse = Nothing
    Set xmlElem = Nothing

    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
