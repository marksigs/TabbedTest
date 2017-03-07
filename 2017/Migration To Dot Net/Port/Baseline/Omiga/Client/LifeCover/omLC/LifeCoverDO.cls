VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "LifeCoverDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      LifeCoverDO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling
'
'Description:   Life Cover Data Object

'Dependencies:
'
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'PSC    26/10/99    Created
'SR     10-12-99    Modified sub LoadLCRateData - class definition according to new table
'                   structuref
'MS     19/07/00    Modified for performance enhancements
'LD     07/11/00    Explicity close recordsets
'LD     07/11/00    Explicity destroy command objects
'AD     08/10/01    Removed the MoveFirst as it's a ForwardOnly cursor. SYS2718
'------------------------------------------------------------------------------------------

Option Explicit
'MS 19/07/00 Modified for performance enhancements
Implements ILifeCoverDO
 
 '=============================================
 'Variable Declaration Section
 '=============================================
 
 '=============================================
 'Constant Declaration Section
 '=============================================

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

    Const strFunctionName As String = "AddDerivedData"

    Dim objErrAssist As New ErrAssist

    AddDerivedData = vstrXMLData
    
    Set objErrAssist = Nothing
    
    Exit Function
            
    Set objErrAssist = Nothing
    
    #If TIME_DO Then
        sngEndTime = Timer
        App.LogEvent TypeName(Me) & "." & strFunctionName & " elapsed " & Format(sngEndTime - sngStartTime, "0.000000"), vbLogEventTypeInformation
    #End If
    
    Exit Function
    
AddDerivedDataVBErr:

    Set objErrAssist = Nothing
    
    AddDerivedData = ""

    Err.Raise Err.Number, Err.Source, Err.Description


End Function
'MS 19/07/00 Modified for performance enhancements
Private Function ILifeCoverDO_GetRates(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode

' header ----------------------------------------------------------------------------------
' description:  Gets the Payment Protection Rate Record
'
' pass:         n/a
'
' return:       Payment Protection Rate Record
'------------------------------------------------------------------------------------------
On Error GoTo ILifeCoverDO_GetRatesVbErr

    Const strFunctionName As String = "ILifeCoverDO_GetRates"

    Dim objXmlAssist As XMLAssist
    Dim objErrAssist As ErrAssist
    Dim objSqlAssist As SQLAssist

    Set objXmlAssist = New XMLAssist
    Set objErrAssist = New ErrAssist
    Set objSqlAssist = New SQLAssist

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()

    Dim strSQL As String
    Dim strXML As String

    Dim objXmlOut As New FreeThreadedDOMDocument40
    objXmlOut.validateOnParse = False
    objXmlOut.setProperty "NewParser", True
    Dim objXmlDoc As New FreeThreadedDOMDocument40
    objXmlDoc.validateOnParse = False
    objXmlDoc.setProperty "NewParser", True
'MS 19/07/00 Modified for performance enhancements
'    Dim objXmlIn As DOMDocument

    Dim objXmlNode As IXMLDOMNode
    Dim objXmlElem As IXMLDOMElement
    Dim objXmlListNode As IXMLDOMNode

    Dim objIADOAssist As IADOAssist
    Dim objIDOAssist As IDOAssist
'MS 19/07/00 Modified for performance enhancements
'    Set objXmlIn = objXmlAssist.load(vstrRequest, TypeName(Me), strFunctionName)

    Dim strCoverType As String
    Dim strApplicantsGender As String
    Dim strApplicantsAge As String
    Dim strTerm As String

    strCoverType = _
        objXmlAssist.GetTagValue(vxmlTableElement, "COVERTYPE")
    strApplicantsGender = _
        objXmlAssist.GetTagValue(vxmlTableElement, "APPLICANTSGENDER")
    strApplicantsAge = _
        objXmlAssist.GetTagValue(vxmlTableElement, "APPLICANTSAGE")
    strTerm = _
        objXmlAssist.GetTagValue(vxmlTableElement, "TERM")

    If Len(strCoverType) = 0 Or Len(strApplicantsGender) = 0 Or _
       Len(strApplicantsAge) = 0 Or Len(strTerm) = 0 Then
        objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeMissingPrimaryTag, _
                                "COVERTYPE = " & strCoverType & _
                                "; APPLICANTSGENDER = " & strApplicantsGender & _
                                "; APPLICANTSAGE = " & strApplicantsAge & _
                                "; TERM = " & strTerm
    End If

    If objContext Is Nothing Then
        Set objIADOAssist = New ADOAssist
        Set objIDOAssist = New DOAssist
    Else
        Set objIADOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If

    strSQL = "SELECT * FROM LIFECOVERRATES LIFECOVER1 WHERE"
    strSQL = strSQL & " LIFECOVER1.LIFECOVERRATESSTARTDATE = "
    strSQL = strSQL & " (SELECT MAX(LIFECOVERRATESSTARTDATE) FROM LIFECOVERRATES WHERE"
    strSQL = strSQL & " LIFECOVERRATESSTARTDATE < = " & objSqlAssist.GetSystemDate
    strSQL = strSQL & " AND COVERTYPE = " & strCoverType
    strSQL = strSQL & " AND APPLICANTSGENDER = "
    strSQL = strSQL & objSqlAssist.FormatString(strApplicantsGender) & ")"
    strSQL = strSQL & " AND LIFECOVER1.HIGHAPPLICANTSAGE = "
    strSQL = strSQL & " (SELECT MIN(HIGHAPPLICANTSAGE) FROM LIFECOVERRATES WHERE"
    strSQL = strSQL & " LIFECOVERRATESSTARTDATE = LIFECOVER1.LIFECOVERRATESSTARTDATE"
    strSQL = strSQL & " AND COVERTYPE = " & strCoverType
    strSQL = strSQL & " AND APPLICANTSGENDER = "
    strSQL = strSQL & objSqlAssist.FormatString(strApplicantsGender)
    strSQL = strSQL & " AND HIGHAPPLICANTSAGE >= " & strApplicantsAge & ")"
    strSQL = strSQL & " AND LIFECOVER1.HIGHTERM = "
    strSQL = strSQL & " (SELECT MIN(HIGHTERM) FROM LIFECOVERRATES WHERE"
    strSQL = strSQL & " LIFECOVERRATESSTARTDATE = LIFECOVER1.LIFECOVERRATESSTARTDATE"
    strSQL = strSQL & " AND COVERTYPE = " & strCoverType
    strSQL = strSQL & " AND APPLICANTSGENDER = "
    strSQL = strSQL & objSqlAssist.FormatString(strApplicantsGender)
    strSQL = strSQL & " AND HIGHTERM >= " & strTerm & ")"
    strSQL = strSQL & " AND LIFECOVER1.COVERTYPE = " & strCoverType
    strSQL = strSQL & " AND LIFECOVER1.APPLICANTSGENDER = "
    strSQL = strSQL & objSqlAssist.FormatString(strApplicantsGender)

    Dim cmd As New ADODB.Command

    cmd.ActiveConnection = objIADOAssist.GetConnStr

    cmd.CommandText = strSQL

    Dim rstThisRecSet As New ADODB.Recordset

    Set rstThisRecSet = cmd.Execute(, , adCmdText)

    If rstThisRecSet.EOF Then
        ' raise application error to be interpreted by calling object
        objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound
    End If

    Dim xmlClassDefDoc As FreeThreadedDOMDocument40

    Dim objIomLCClassDef As IomLCClassDef
'MS 19/07/00 Modified for performance enhancements
'    Dim strClassDef As String
'    strClassDef = LoadLCRateData()
    Set objIomLCClassDef = New omLCClassDef
    ' AD 08/10/01. Removed the MoveFirst as it's a ForwardOnly cursor. SYS2718
    
    Dim vxmlTempElement As IXMLDOMElement

    Set xmlClassDefDoc = objIomLCClassDef.LoadLifeCoverRatesData()
    
    ' loop through the record set
    While Not rstThisRecSet.EOF
'MS 19/07/00 Modified for performance enhancements
 '       strXML = objIDOAssist.GetXMLFromRecordSet(rstThisRecSet, strClassDef)
        Set vxmlTempElement = objIDOAssist.GetXMLFromRecordSet(rstThisRecSet, xmlClassDefDoc)
'MS 19/07/00 Modified for performance enhancements
'        Set objXmlDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
'        Set objXmlNode = objXmlOut.appendChild(objXmlDoc.documentElement)
        objXmlOut.appendChild vxmlTempElement
        rstThisRecSet.MoveNext
    Wend
        rstThisRecSet.Close

    Set rstThisRecSet = Nothing
    Set cmd = Nothing

    Set ILifeCoverDO_GetRates = objXmlOut.documentElement

    Set objErrAssist = Nothing
    Set objXmlDoc = Nothing
    Set objXmlOut = Nothing
    Set objIADOAssist = Nothing
    Set objXmlAssist = Nothing
'MS 19/07/00 Modified for performance enhancements
'    Set objXmlIn = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ILifeCoverDO_GetRatesVbErr:

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
    Set objIADOAssist = Nothing
    Set objIDOAssist = Nothing
    Set objXmlAssist = Nothing
'MS 19/07/00 Modified for performance enhancements
'    Set objXmlIn = Nothing

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


