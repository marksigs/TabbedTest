VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "PaymentProtectionDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      PaymentProtectionDO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling
'
'Description:   Payment Protection Data Object

'Dependencies:
'
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date     Description
'PSC    14/10/99 Created
'SR     08/12/99 Modified method GetRates
'LD     07/11/00 Explicity close recordsets
'LD     07/11/00 Explicity destroy command objects
'AD     10/10/01 Removed the MoveFirst as it's a ForwardOnly cursor. SYS2718.
'------------------------------------------------------------------------------------------

Option Explicit


 
 '=============================================
 'Variable Declaration Section
 '=============================================
 
 '=============================================
 'Constant Declaration Section
 '=============================================
Private Function LoadPPProductData()
'------------------------------------------------------------------------------------------
'
'
'
' *****************************************************************************
Dim XMLstr As String

    Dim strXML As String
    
    strXML = strXML & "<TABLENAME>"
    strXML = strXML & "PAYMENTPROTECTIONPRODUCT"
    strXML = strXML & "<PRIMARYKEY>PPPRODUCTCODE<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>PPSTARTDATE<TYPE>dbdtDate</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>PPENDDATE<TYPE>dbdtDate</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>PRODUCTNAME<TYPE>dbdtString</TYPE></OTHERS>"
    ' APS 03/03/00 - Field removed from table
    'strXML = strXML & "<OTHERS>MINAPPLICANTAGE<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>MAXAPPLICANTAGE<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>PAYMENTPROTECTIONRATESNUMBER<TYPE>dbdtLong</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>DISPLAYORDER<TYPE>dbdtInt</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"

    LoadPPProductData = strXML

End Function

Private Function LoadPPRateData()
'------------------------------------------------------------------------------------------
'
'
'
' *****************************************************************************
Dim XMLstr As String

    Dim strXML As String
    
    strXML = strXML & "<TABLENAME>"
    strXML = strXML & "PAYMENTPROTECTIONRATES"
    strXML = strXML & "<PRIMARYKEY>PAYMENTPROTECTIONRATESNUMBER<TYPE>dbdtLong</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>PPSTARTDATE<TYPE>dbdtDate</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>CHANNELID<TYPE>dbdtString</TYPE></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>APPLICANTSGENDER<TYPE>dbdtComboId</TYPE>"
    strXML = strXML & "<COMBO>Sex</COMBO></PRIMARYKEY>"
    strXML = strXML & "<PRIMARYKEY>HIGHAPPLICANTSAGE<TYPE>dbdtInt</TYPE></PRIMARYKEY>"
    strXML = strXML & "<OTHERS>ASURATE<TYPE>dbdtDouble</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>ASRATE<TYPE>dbdtDouble</TYPE></OTHERS>"
    strXML = strXML & "<OTHERS>URATE<TYPE>dbdtDouble</TYPE></OTHERS>"
    strXML = strXML & "</TABLENAME>"

    LoadPPRateData = strXML

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

Public Function FindProductList() As String
' header ----------------------------------------------------------------------------------
' description:  Gets a list of all active payment protection products
'
' pass:         n/a
'
' return:       List of active products
'------------------------------------------------------------------------------------------
On Error GoTo FindProductListVbErr
    
    Const strFunctionName As String = "FindProductList"
    
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim objSQLAssist As New SQLAssist
    
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
    
    Dim objXmlNode As IXMLDOMNode
    Dim objXmlElem As IXMLDOMElement
    Dim objXmlListNode As IXMLDOMNode
    
    Dim objADOAssist As ADOAssist
    Dim objDOAssist As DOAssist
    Dim strTable As String
    strTable = "PAYMENTPROTECTIONPRODUCT"
    
    If objContext Is Nothing Then
        Set objADOAssist = New ADOAssist
        Set objDOAssist = New DOAssist
    Else
        Set objADOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
        Set objDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    strSQL = "SELECT * FROM PAYMENTPROTECTIONPRODUCT WHERE"
    strSQL = strSQL & " PPSTARTDATE <= " & objSQLAssist.GetSystemDate
    strSQL = strSQL & " AND (PPENDDATE >= " & objSQLAssist.GetSystemDate
    strSQL = strSQL & " OR PPENDDATE IS NULL) ORDER BY DISPLAYORDER"
          
    Dim cmd As New ADODB.Command
        
    cmd.ActiveConnection = objADOAssist.GetConnStr
    
    cmd.CommandText = strSQL
    
    Dim rstThisRecSet As New ADODB.Recordset
        
    Set rstThisRecSet = cmd.Execute(, , adCmdText)
                
    If rstThisRecSet.EOF Then
        ' raise application error to be interpreted by calling object
        objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound
    End If
        
    Set objXmlElem = objXmlOut.createElement(strTable & "LIST")
    Set objXmlListNode = objXmlOut.appendChild(objXmlElem)
    
    Dim strClassDef As String
    strClassDef = LoadPPProductData()
    
        
    ' AD 10/10/01 Removed the MoveFirst as it's a ForwardOnly cursor. SYS2718.
    
    ' loop through the record set
    While Not rstThisRecSet.EOF
        strXML = objDOAssist.GetXMLFromRecordSet(rstThisRecSet, strClassDef)
        Set objXmlDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
        Set objXmlNode = objXmlListNode.appendChild(objXmlDoc.documentElement)
        rstThisRecSet.MoveNext
    Wend
    rstThisRecSet.Close

    Set rstThisRecSet = Nothing
    Set cmd = Nothing
    
    FindProductList = objXmlOut.xml
    
    Set objErrAssist = Nothing
    Set objXmlDoc = Nothing
    Set objXmlOut = Nothing
    Set objADOAssist = Nothing
    Set objXmlAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
            
    Exit Function

FindProductListVbErr:
    
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

Public Function GetRates(ByVal vstrRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Gets the Payment Protection Rate Record
'
' pass:         n/a
'
' return:       Payment Protection Rate Record
'------------------------------------------------------------------------------------------
On Error GoTo GetRatesVbErr
    
    Const strFunctionName As String = "GetRates"
    
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim objSQLAssist As New SQLAssist
    
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
    Dim objXmlIn As FreeThreadedDOMDocument40
    
    Dim objXmlNode As IXMLDOMNode
    Dim objXmlElem As IXMLDOMElement
    Dim objXmlListNode As IXMLDOMNode
    
    Dim objADOAssist As ADOAssist
    Dim objDOAssist As DOAssist
    
    Set objXmlIn = objXmlAssist.load(vstrRequest, TypeName(Me), strFunctionName)
    
    Dim strPPRateNumber As String
    Dim strChannelId As String
    Dim strApplicantsGender As String
    Dim strApplicantsAge As String
    
    strPPRateNumber = _
        objXmlAssist.GetTagValue(objXmlIn.documentElement, "PAYMENTPROTECTIONRATESNUMBER")
    strChannelId = _
        objXmlAssist.GetTagValue(objXmlIn.documentElement, "CHANNEL")
    strApplicantsGender = _
        objXmlAssist.GetTagValue(objXmlIn.documentElement, "APPLICANTSGENDER")
    strApplicantsAge = _
        objXmlAssist.GetTagValue(objXmlIn.documentElement, "APPLICANTSAGE")
    
    If Len(strPPRateNumber) = 0 Or Len(strChannelId) = 0 Or _
       Len(strApplicantsGender) = 0 Or Len(strApplicantsAge) = 0 Then
        objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeMissingPrimaryTag, _
                                "PAYMENTPROTECTIONRATESNUMBER = " & strPPRateNumber & _
                                "; CHANNELINDICATOR = " & strChannelId & _
                                "; APPLICANTSGENDER = " & strApplicantsGender & _
                                "; APPLICANTSAGE = " & strApplicantsAge
    End If
    
    If objContext Is Nothing Then
        Set objADOAssist = New ADOAssist
        Set objDOAssist = New DOAssist
    Else
        Set objADOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
        Set objDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    'Lowest of HighApplicantsAge that is NOT less than ApplicantsAge(from Input)
    Dim strSQLLowestHighApplicantsAge As String
    
    'Highest of PPStartDate that is below System Date (recently started PPRate)
    Dim strSQLHighestPPStartDate As String
    
    strSQLLowestHighApplicantsAge = _
        " (SELECT MIN(HIGHAPPLICANTSAGE) FROM PAYMENTPROTECTIONRATES PPRATES2 " _
        & " WHERE PPRATES2.HIGHAPPLICANTSAGE >= " & strApplicantsAge _
        & " AND PPRATES2.PAYMENTPROTECTIONRATESNUMBER = " & strPPRateNumber _
        & " AND PPRATES2.PPSTARTDATE = PPRATES1.PPSTARTDATE " _
        & " AND PPRATES2.CHANNELID = " & objSQLAssist.FormatString(strChannelId) _
        & " AND PPRATES2.APPLICANTSGENDER = " & strApplicantsGender & ")"
    
    strSQLHighestPPStartDate = _
           " (SELECT MAX(PPSTARTDATE) FROM PAYMENTPROTECTIONRATES WHERE " _
           & " PPSTARTDATE <= " & objSQLAssist.GetSystemDate _
           & " AND PAYMENTPROTECTIONRATESNUMBER = " & strPPRateNumber _
           & " AND CHANNELID = " & objSQLAssist.FormatString(strChannelId) _
           & " AND APPLICANTSGENDER = " & strApplicantsGender & ")"
    
    ' SQL used to get the data from the database
    strSQL = "SELECT * FROM PAYMENTPROTECTIONRATES PPRATES1 WHERE"
    strSQL = strSQL & " PPRATES1.PAYMENTPROTECTIONRATESNUMBER = " & strPPRateNumber
    strSQL = strSQL & " AND PPRATES1.PPSTARTDATE = " & strSQLHighestPPStartDate
    strSQL = strSQL & " AND CHANNELID = " & objSQLAssist.FormatString(strChannelId)
    strSQL = strSQL & " AND APPLICANTSGENDER = " & strApplicantsGender
    strSQL = strSQL & " AND HIGHAPPLICANTSAGE = " & strSQLLowestHighApplicantsAge
    
    
    Dim cmd As New ADODB.Command
        
    cmd.ActiveConnection = objADOAssist.GetConnStr
    
    cmd.CommandText = strSQL
    
    Dim rstThisRecSet As New ADODB.Recordset
        
    Set rstThisRecSet = cmd.Execute(, , adCmdText)
                
    If rstThisRecSet.EOF Then
        ' raise application error to be interpreted by calling object
        objErrAssist.ThrowError TypeName(Me), strFunctionName, 232
    End If
        
    Dim strClassDef As String
    strClassDef = LoadPPRateData()
    
    ' AD 10/10/01 Removed the MoveFirst as it's a ForwardOnly cursor. SYS2718.
    
    ' loop through the record set
    While Not rstThisRecSet.EOF
        strXML = objDOAssist.GetXMLFromRecordSet(rstThisRecSet, strClassDef)
        Set objXmlDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
        Set objXmlNode = objXmlOut.appendChild(objXmlDoc.documentElement)
        rstThisRecSet.MoveNext
    Wend
    rstThisRecSet.Close

    Set rstThisRecSet = Nothing
    Set cmd = Nothing
    
    GetRates = objXmlOut.xml
    
    Set objErrAssist = Nothing
    Set objXmlDoc = Nothing
    Set objXmlOut = Nothing
    Set objADOAssist = Nothing
    Set objXmlAssist = Nothing
    Set objXmlIn = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
            
    Exit Function

GetRatesVbErr:
    
    ' Close the recordset if still open
        If Not rstThisRecSet Is Nothing Then
                If rstThisRecSet.State = adStateOpen Then
                        rstThisRecSet.Close
                End If
        End If

    Set rstThisRecSet = Nothing
    Set objXmlDoc = Nothing
    Set objXmlOut = Nothing
    Set objADOAssist = Nothing
    Set objXmlAssist = Nothing
    Set objXmlIn = Nothing
        
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


