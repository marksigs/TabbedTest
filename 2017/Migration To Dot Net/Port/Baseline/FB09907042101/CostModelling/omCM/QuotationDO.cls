VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "QuotationDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'Workfile:      QuotationDO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling

'Description:   Quotation Data Object
'               contains (implied) standard interface:

'               Public Sub Update(vstrXMLRequest As String)

'Dependencies:
'               Add any other dependent components
'
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'PSC    09/11/1999  Created
'MCS    05/01/2000  Updated to use interfaces
'PSC    11/01/2000  Add Create, FindList, GetQuotationSummary, StoredQuoteFindList
'                   CountQuotationsForMortgageSubQuote
'AY     21/03/00    Life subquote details no longer returned unnecessarily on GetQuotationSummary
'MS     04/07/00    SYS1048 GetAcceptedBCSubQuoteNumber Method Added
'IVW    09/11/00    Amended to disconnect record sets more efficiently
'PSC    05/03/01    SYS1879 Added GetQuotationLoanSummary
'AD     05/10/01    Removed MoveFirst as it's a ForwardOnly cursor. SYS2718.
'JLD    15/01/02    SYS3499 error when closing connection
'------------------------------------------------------------------------------------------
'BMIDS Specific History:
'
'Prog   Date        AQR         Description
'LD     02/01/2003  BM0391      Use COUNT(1) rather than COUNT(*)
'GHun   22/10/2003  BMIDS624    Added HaveRatesChanged method
'HMA    14/10/2004  BM0464      Changed GetQuotationLoanSummary
'------------------------------------------------------------------------------------------
'MARS History:
'
'Prog   Date        Description
'MV     14/09/2005  MAR35 - New IQuotationDO_GetAllQuotationsForApplication()
'MV     22/10/2005  MAR270 - AMended IQuotationDO_GetQuotationSummary() Removed LifeBenifSequenceNumber
'Maha T 25/11/2005  MAR632 - AMended IQuotationDO_GetQuotationSummary() Don't require anymore to check
'                            LIFE SUBQUOTE. (Code Commented)
'HMA    28/03/2006  MAR1500  Changed HaveRatesChanged to accept Quotation Number
'GHun   10/07/2006  MAR1893 Commented out remaining LifeSubQuote functionality as it is redundant
'------------------------------------------------------------------------------------------
Option Explicit
Implements IQuotationDO
Implements ObjectControl

Private m_objContext As ObjectContext
Private m_objErrAssist As ErrAssist
Private m_objXmlAssist As XMLAssist
Private m_objIDOAssist As IDOAssist
Private m_objSQLAssist As SQLAssist

Private Function IQuotationDO_CountQuotationsForMortgageSubQuote(ByVal vxmlTableElement As IXMLDOMElement) As Integer
    
' header ----------------------------------------------------------------------------------
' description: Counts the number of quotations a mortgage sub quote is attached to
' pass:
'   vxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'
'
'   Input XML
'
'   <QUOTATION>
'       <APPLICATIONNUMBER>Value</APPLICATIONNUMBER>
'       <APPLICATIONFACTFINDNUMBER>Value</APPLICATIONFACTFINDNUMBER>
'        <MORTGAGESUBQUOTENUMBER>Value</MORTGAGESUBQUOTENUMBER>
'   </QUOTATION>
'------------------------------------------------------------------------------------------
On Error GoTo CountQuotationsForMortgageSubQuoteVbErr

    Const strFunctionName As String = "IQuotationDO_CountQuotationsForMortgageSubQuote"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
        
    Dim objErrAssist As New ErrAssist
    Dim objIDOAssist As IDOAssist
    Dim objIADOAssist As IADOAssist
    Dim objXmlAssist As New XMLAssist
    Dim objSQLAssist As New SQLAssist
                 
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
        Set objIADOAssist = New ADOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
        Set objIADOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    End If
     
    Dim strApplicationNo As String
    Dim strApplicationFFNo As String
    Dim strMortgageSubQuoteNo As String
    
    strApplicationNo = objXmlAssist.GetTagValue(vxmlTableElement, "APPLICATIONNUMBER")
    strApplicationFFNo = objXmlAssist.GetTagValue(vxmlTableElement, _
                                                  "APPLICATIONFACTFINDNUMBER")
    strMortgageSubQuoteNo = objXmlAssist.GetTagValue(vxmlTableElement, _
                                                     "MORTGAGESUBQUOTENUMBER")
                                                     
    If Len(strApplicationNo) = 0 Or Len(strApplicationFFNo) = 0 Or _
       Len(strMortgageSubQuoteNo) = 0 Then
        Dim strDetails As String
        
        strDetails = "APPLICATIONNUMBER = " & strApplicationNo
        strDetails = strDetails & "; APPLICATIONFACTFINDNUMBER = " & strApplicationFFNo
        strDetails = strDetails & "; MORTGAGESUBQUOTENUMBER = " & strMortgageSubQuoteNo
        
        objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeMissingPrimaryTag, _
                                strDetails
    End If
    
    Dim strSQL As String
    
    strSQL = "SELECT COUNT(1) FROM QUOTATION WHERE APPLICATIONNUMBER = "
    strSQL = strSQL & objSQLAssist.FormatString(strApplicationNo)
    strSQL = strSQL & " AND APPLICATIONFACTFINDNUMBER = " & strApplicationFFNo
    strSQL = strSQL & " AND MORTGAGESUBQUOTENUMBER = " & strMortgageSubQuoteNo
    
    'IVW - Disconnecting record sets
    
    Dim cmd As ADODB.Command
    Dim conn As ADODB.Connection
    Dim rstThisRecSet As ADODB.Recordset
    
    Set cmd = New ADODB.Command
    Set conn = New ADODB.Connection
    Set rstThisRecSet = New ADODB.Recordset

    cmd.CommandText = strSQL
    
    rstThisRecSet.CursorLocation = adUseClient
    rstThisRecSet.CursorType = adOpenForwardOnly
    rstThisRecSet.LockType = adLockReadOnly
    
    conn.Open objIADOAssist.GetConnStr
    cmd.ActiveConnection = conn
            
    Set rstThisRecSet.Source = cmd
    
    rstThisRecSet.Open
    
    Set rstThisRecSet.ActiveConnection = Nothing
    
    
 
    'IVW
    
    IQuotationDO_CountQuotationsForMortgageSubQuote = rstThisRecSet.Fields(0).Value
    
    Set objErrAssist = Nothing
    Set objIDOAssist = Nothing
    Set objIADOAssist = Nothing
    Set objSQLAssist = Nothing
    Set cmd = Nothing
        rstThisRecSet.Close
        Set rstThisRecSet = Nothing
        conn.Close
    Set conn = Nothing
    
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

CountQuotationsForMortgageSubQuoteVbErr:
    
    ' Close the recordset if still open
        If Not rstThisRecSet Is Nothing Then
                If rstThisRecSet.State = adStateOpen Then
                        rstThisRecSet.Close
                End If
        End If


    ' Close the connection if still open
    If conn.State = adStateOpen Then
        conn.Close
    End If
    
    Set objIDOAssist = Nothing
    Set objIADOAssist = Nothing
    Set objSQLAssist = Nothing
    Set cmd = Nothing
    Set conn = Nothing
    Set rstThisRecSet = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
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
    Err.Raise Err.number, Err.Source, Err.Description

End Function

Private Function IQuotationDO_GetAllQuotationsForApplication(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
    
    On Error GoTo GetAllQuotationsForApplicationVbErr

    Const strFunctionName As String = "IQuotationDO_GetAllQuotationsForApplication"
    
    Dim objIDOAssist As IDOAssist
    
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIomCMClassDef As IomCMClassDef
    Dim strApplicationNumber  As String
    Dim strApplicationFactFindNumber  As String
    Dim strSQL As String
    Dim cmd As ADODB.Command
    Dim conn As ADODB.Connection
    Dim rstThisRecSet As ADODB.Recordset
    Dim xmlElement  As IXMLDOMElement
    Dim xmlQuotationNode  As IXMLDOMNode
       
    Dim xmlOut As FreeThreadedDOMDocument40
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    
    Dim xmlResponseElem  As IXMLDOMElement
    
    Set xmlResponseElem = xmlOut.createElement("QUOTATIONLIST")
    
    If m_objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set xmlElement = m_objXmlAssist.GetMandatoryNode(vxmlTableElement, "//QUOTATION")
    
    strApplicationNumber = m_objXmlAssist.GetTagValue(xmlElement, "APPLICATIONNUMBER")
    strApplicationFactFindNumber = m_objXmlAssist.GetTagValue(xmlElement, "APPLICATIONFACTFINDNUMBER")
                   
    strSQL = "SELECT * FROM GETALLQUOTATIONSFORAPPLICATION WHERE APPLICATIONNUMBER = "
    strSQL = strSQL & m_objSQLAssist.FormatString(strApplicationNumber)
    strSQL = strSQL & " AND APPLICATIONFACTFINDNUMBER = " & strApplicationFactFindNumber
    strSQL = strSQL & " ORDER BY QUOTATIONNUMBER"
            
    Set cmd = New ADODB.Command
    Set conn = New ADODB.Connection
    Set rstThisRecSet = New ADODB.Recordset

    cmd.CommandText = strSQL
    
    rstThisRecSet.CursorLocation = adUseClient
    rstThisRecSet.CursorType = adOpenForwardOnly
    rstThisRecSet.LockType = adLockReadOnly
    
    conn.Open adoGetDbConnectString
    cmd.ActiveConnection = conn
            
    Set rstThisRecSet.Source = cmd
    
    rstThisRecSet.Open
    
    Set rstThisRecSet.ActiveConnection = Nothing
    
    Set objIomCMClassDef = New omCMClassDef
    Set xmlClassDefDoc = objIomCMClassDef.LoadGetAllQuotationsForAppViewData()
    
    Do While Not rstThisRecSet.EOF
        Set xmlQuotationNode = objIDOAssist.GetXMLFromRecordSet(rstThisRecSet, xmlClassDefDoc)
        xmlResponseElem.appendChild xmlQuotationNode
        rstThisRecSet.MoveNext
    Loop
    
    Set IQuotationDO_GetAllQuotationsForApplication = xmlResponseElem
    
    Set objIomCMClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Exit Function

GetAllQuotationsForApplicationVbErr:
    
    Set objIomCMClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If
    
    If Not m_objContext Is Nothing Then
        If m_objErrAssist.IsSystemError Then
            m_objContext.SetAbort
        Else
            m_objContext.SetComplete
        End If
    End If
    
    Set m_objErrAssist = Nothing
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.number, Err.Source, Err.Description
End Function



Private Function IQuotationDO_Create(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   create an instance of the persistant data associated with this data object
'   for each set of data in the request
' pass:
'   vxmlTableElement     xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo CreateVbErr

    Const strFunctionName As String = "IQuotationDO_Create"
   
    Dim objErrAssist As New ErrAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIomCMClassDef As IomCMClassDef
    Dim xmlElement As IXMLDOMElement
    
    Dim objIDOAssist As IDOAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
                    
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIomCMClassDef = New omCMClassDef
    Set xmlClassDefDoc = objIomCMClassDef.LoadQuotationData()
    
    ' Set up the next sequence number
    objIDOAssist.GenerateSequenceNumber vxmlTableElement, xmlClassDefDoc, "QUOTATIONNUMBER"
    
    ' Create new quotation
    objIDOAssist.Create vxmlTableElement, xmlClassDefDoc
    
    Set xmlElement = vxmlTableElement.getElementsByTagName("QUOTATIONNUMBER").Item(0)
    
    Set IQuotationDO_Create = xmlElement.cloneNode(True)
    
    Set objIomCMClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

CreateVbErr:
    
    Set objIomCMClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
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
    Err.Raise Err.number, Err.Source, Err.Description

End Function

Private Function IQuotationDO_FindList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for all instances of the persistant data associated with
'   this data object for the values supplied
' pass:
'   vxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo FindListVbErr

    Const strFunctionName As String = "IQuotationDO_FindList"

    Dim objErrAssist As New ErrAssist
    
    Dim objIDOAssist As IDOAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIomCMClassDef As IomCMClassDef
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
         
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
     
    Set objIomCMClassDef = New omCMClassDef
    Set xmlClassDefDoc = objIomCMClassDef.LoadQuotationData()
    Set IQuotationDO_FindList = objIDOAssist.FindList(vxmlTableElement, xmlClassDefDoc)
    
    Set objIomCMClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objErrAssist = Nothing
    Set objIDOAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

FindListVbErr:
    
    Set objIomCMClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
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
    Err.Raise Err.number, Err.Source, Err.Description

End Function

Private Function IQuotationDO_GetData(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   this data object
' pass:
'   vxmlTableElement     xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo GetDataVbErr

    Const strFunctionName As String = "IQuotationDO_GetData"
    
    Dim objErrAssist As New ErrAssist
    
    Dim objIDOAssist As IDOAssist
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIomCMClassDef As IomCMClassDef
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
            
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIomCMClassDef = New omCMClassDef
    Set xmlClassDefDoc = objIomCMClassDef.LoadQuotationData()
    Set IQuotationDO_GetData = objIDOAssist.GetData(vxmlTableElement, xmlClassDefDoc)
    
    Set objIomCMClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

GetDataVbErr:
    
    Set objIomCMClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
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
    Err.Raise Err.number, Err.Source, Err.Description
End Function

Private Function IQuotationDO_GetQuotationDetails(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   this data object
' pass:
'   vxmlRequest     xml Request data stream containing data to which identifies
'                   the instance of the persistant data to be retrieved
'   vxmlClassDef    xml Class Def to parse vxmlRequest against
' return:
'   GetDataEx       string containing XML data stream representation of
'                   data retrieved
' Raise Errors: if record not found, raise omiga4RecordNotFound
'------------------------------------------------------------------------------------------
On Error GoTo GetQuotationDetailsVbErr

    Const strFunctionName As String = "IQuotationDO_GetQuotationDetails"

    Dim objErrAssist As ErrAssist
    Set objErrAssist = New ErrAssist
        
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
        
    Dim objIADOAssist As IADOAssist
    Dim objIDOAssist As IDOAssist
    Dim objSQLAssist As SQLAssist
    Dim xmlElement As IXMLDOMElement
    Dim objXmlAssist As XMLAssist
    Dim xmlQuoteDefDoc As FreeThreadedDOMDocument40
    Dim xmlMortSubQuoteDefDoc As FreeThreadedDOMDocument40
    'Dim xmlLifeSubQuoteDefDoc As FreeThreadedDOMDocument40
    Dim xmlPPSubQuoteDefDoc As FreeThreadedDOMDocument40
    Dim xmlBCSubQuoteDefDoc As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlQuotationElement As IXMLDOMElement
    Dim objIomCMClassDef As IomCMClassDef
    
    If objContext Is Nothing Then
        Set objIADOAssist = New ADOAssist
        Set objIDOAssist = New DOAssist
    Else
        Set objIADOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
                
    Set xmlOut = xmlCreateDOMObject()
    
    If vXMLRequest.nodeName = "QUOTATION" Then
       Set xmlElement = vXMLRequest
    Else
        Set xmlElement = vXMLRequest.getElementsByTagName("QUOTATION").Item(0)
    End If
    
    If xmlElement Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "QUOTATION tag not found"
    End If
    
    Dim strInApplicationNumber As String
    Dim strInApplicationFactFindNumber As String
    Dim strInQuotationNumber As String

    Set objXmlAssist = New XMLAssist

    'Need to get the values for these fields out of the XML passed in
    strInApplicationNumber = objXmlAssist.GetTagValue(xmlElement, "APPLICATIONNUMBER")
    strInApplicationFactFindNumber = objXmlAssist.GetTagValue(xmlElement, "APPLICATIONFACTFINDNUMBER")
'    strInQuotationNumber = objXmlAssist.GetTagValue(xmlElement, "QUOTATIONNUMBER")
    strInQuotationNumber = objXmlAssist.GetTagValue(xmlElement, "ACTIVEQUOTENUMBER")

    If Len(strInApplicationNumber) = 0 Or _
        Len(strInApplicationFactFindNumber) = 0 Or _
        Len(strInQuotationNumber) = 0 Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingPrimaryTag, _
        "APPLICATIONNUMBER =" & strInApplicationNumber & _
        "; APPLICATIONFACTFINDNUMBER =" & strInApplicationFactFindNumber & _
        "; ACTIVEQUOTENUMBER =" & strInQuotationNumber
'        "; QUOTATIONNUMBER =" & strInQuotationNumber

    End If
                                
    Set objSQLAssist = New SQLAssist
                                
    Dim strSQL As String
                   
    strSQL = "SELECT * FROM GETQUOTATIONDETAILS WHERE APPLICATIONNUMBER = "
    strSQL = strSQL & objSQLAssist.FormatString(strInApplicationNumber)
    strSQL = strSQL & " AND APPLICATIONFACTFINDNUMBER = " & strInApplicationFactFindNumber
    strSQL = strSQL & " AND QUOTATIONNUMBER = " & strInQuotationNumber
            
    'IVW - Disconnecting record sets
        
    Dim cmd As ADODB.Command
    Dim conn As ADODB.Connection
    Dim rstThisRecSet As ADODB.Recordset
    
    Set cmd = New ADODB.Command
    Set conn = New ADODB.Connection
    Set rstThisRecSet = New ADODB.Recordset

    cmd.CommandText = strSQL
    
    rstThisRecSet.CursorLocation = adUseClient
    rstThisRecSet.CursorType = adOpenForwardOnly
    rstThisRecSet.LockType = adLockReadOnly
    
    conn.Open objIADOAssist.GetConnStr
    cmd.ActiveConnection = conn
            
    Set rstThisRecSet.Source = cmd
    
    rstThisRecSet.Open
    
    Set rstThisRecSet.ActiveConnection = Nothing
    
    
    'IVW
        
    Set objIomCMClassDef = New omCMClassDef
    
    Set xmlQuoteDefDoc = objIomCMClassDef.LoadQuotationData(QuotationView)
    Set xmlMortSubQuoteDefDoc = objIomCMClassDef.LoadMortgageSubQuoteData(MortgageSubQuoteView)
    'Set xmlLifeSubQuoteDefDoc = objIomCMClassDef.LoadLifeSubQuoteData(LifeSubQuoteView)
    Set xmlBCSubQuoteDefDoc = objIomCMClassDef.LoadBCSubQuoteData(BCSubQuoteView)
    Set xmlPPSubQuoteDefDoc = objIomCMClassDef.LoadPPSubQuoteData(PPSubQuoteView)
                       
    If Not rstThisRecSet.EOF Then
        
        Set xmlQuotationElement = objIDOAssist.GetXMLFromRecordSet(rstThisRecSet, xmlQuoteDefDoc)
        xmlOut.appendChild xmlQuotationElement
        
        'Add Mortgage Sub Quote if present
        If Not IsNull(rstThisRecSet("MORTGAGESUBQUOTENUMBER").Value) Then
           objIDOAssist.GetXMLFromRecordSet rstThisRecSet, _
                                            xmlMortSubQuoteDefDoc, _
                                            xmlQuotationElement
        End If
    
        'MAR1893 GHun Redundant
        ''Add Life Sub Quote if present
        'If Not IsNull(rstThisRecSet("LIFESUBQUOTENUMBER").Value) Then
        '   objIDOAssist.GetXMLFromRecordSet rstThisRecSet, _
        '                                    xmlLifeSubQuoteDefDoc, _
        '                                    xmlQuotationElement
        '
        'End If
        'MAR1893 End
        
        'Add B & C Sub Quote if present
        If Not IsNull(rstThisRecSet("BCSUBQUOTENUMBER").Value) Then
           objIDOAssist.GetXMLFromRecordSet rstThisRecSet, _
                                            xmlBCSubQuoteDefDoc, _
                                            xmlQuotationElement
        
        End If
        
        'Add PP Sub Quote if present
        If Not IsNull(rstThisRecSet("PPSUBQUOTENUMBER").Value) Then
           objIDOAssist.GetXMLFromRecordSet rstThisRecSet, _
                                            xmlPPSubQuoteDefDoc, _
                                            xmlQuotationElement
        
        End If
                
    Else
        ' raise application error to be interpreted by calling object
        objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotfound
    End If
        
    Set IQuotationDO_GetQuotationDetails = xmlQuotationElement
    
    Set objErrAssist = Nothing
    Set cmd = Nothing
    rstThisRecSet.Close
    Set rstThisRecSet = Nothing
    conn.Close
    Set conn = Nothing
    
    Set objIADOAssist = Nothing
    Set objIDOAssist = Nothing
    Set objSQLAssist = Nothing
    Set xmlElement = Nothing
    Set objXmlAssist = Nothing
    Set xmlQuoteDefDoc = Nothing
    Set xmlMortSubQuoteDefDoc = Nothing
    'Set xmlLifeSubQuoteDefDoc = Nothing
    Set xmlPPSubQuoteDefDoc = Nothing
    Set xmlBCSubQuoteDefDoc = Nothing
    Set xmlOut = Nothing
    Set xmlQuotationElement = Nothing
    Set objIomCMClassDef = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Set objContext = Nothing

    Exit Function

GetQuotationDetailsVbErr:
    

    Set objIADOAssist = Nothing
    Set objIDOAssist = Nothing
    Set objSQLAssist = Nothing
    Set xmlElement = Nothing
    Set objXmlAssist = Nothing
    Set xmlQuoteDefDoc = Nothing
    Set xmlMortSubQuoteDefDoc = Nothing
    'Set xmlLifeSubQuoteDefDoc = Nothing
    Set xmlPPSubQuoteDefDoc = Nothing
    Set xmlBCSubQuoteDefDoc = Nothing
    Set xmlOut = Nothing
    Set xmlQuotationElement = Nothing
    Set objIomCMClassDef = Nothing
    
    ' Close the recordset if still open
    If Not rstThisRecSet Is Nothing Then
        If rstThisRecSet.State = adStateOpen Then
            rstThisRecSet.Close
        End If
    End If

    ' Close the connection if still open
    If conn.State = adStateOpen Then
        conn.Close
    End If
    
    Set cmd = Nothing
    Set conn = Nothing
    Set rstThisRecSet = Nothing

    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If
    
    If Not objContext Is Nothing Then
        If objErrAssist.IsSystemError Then
            objContext.SetAbort
        Else
            objContext.SetComplete
        End If
    End If
    
    Set objErrAssist = Nothing
    Set objContext = Nothing
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.number, Err.Source, Err.Description
End Function

Private Function IQuotationDO_GetQuotationSummary(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   this data object
' pass:
'   vxmlRequest     xml Request data stream containing data to which identifies
'                   the instance of the persistant data to be retrieved
'   vxmlClassDef    xml Class Def to parse vxmlRequest against
' return:
'   GetDataEx       string containing XML data stream representation of
'                   data retrieved
' Raise Errors: if record not found, raise omiga4RecordNotFound
'------------------------------------------------------------------------------------------
On Error GoTo GetQuotationSummaryVbErr

    Const strFunctionName As String = "IQuotationDO_GetQuotationSummary"

    Dim objErrAssist As ErrAssist
    Set objErrAssist = New ErrAssist
        
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
        
    Dim objIADOAssist As IADOAssist
    Dim objIDOAssist As IDOAssist
    Dim objSQLAssist As SQLAssist
    Dim xmlElement As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    Dim objXmlAssist As XMLAssist
    Dim xmlQuotationSummaryDefDoc As FreeThreadedDOMDocument40
    Dim xmlLifeBenefitDefDoc As FreeThreadedDOMDocument40
    Dim xmlLoanComponentDefDoc As FreeThreadedDOMDocument40
    Dim xmlMortProdLangDefDoc As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlQuoteSummaryElem As IXMLDOMElement
    Dim objIomCMClassDef As IomCMClassDef
    Dim objIomMPClassDef As IomMPClassDef
    Dim xmlMortProdLangTableElem As IXMLDOMElement
    Dim xmlLoanCompTableElem As IXMLDOMElement
    Dim xmlLoanCompListElem As IXMLDOMElement
    Dim xmlLifeBenefitListElem As IXMLDOMElement
    Dim fldFields As ADODB.Fields
    
    Set xmlOut = xmlCreateDOMObject()
    
    If objContext Is Nothing Then
        Set objIADOAssist = New ADOAssist
        Set objIDOAssist = New DOAssist
    Else
        Set objIADOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    If vXMLRequest.nodeName = "QUOTATION" Then
       Set xmlElement = vXMLRequest
    Else
        Set xmlElement = vXMLRequest.getElementsByTagName("QUOTATION").Item(0)
    End If
    
    If xmlElement Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "QUOTATION tag not found"
    End If
    
    Dim strInApplicationNumber As String
    Dim strInApplicationFactFindNumber As String
    Dim strInQuotationNumber As String
    Dim strCountryOfOrigin As String
    Dim strLanguage As String

    Set objXmlAssist = New XMLAssist

    'Need to get the values for these fields out of the XML passed in
    strInApplicationNumber = objXmlAssist.GetTagValue(xmlElement, "APPLICATIONNUMBER")
    strInApplicationFactFindNumber = objXmlAssist.GetTagValue(xmlElement, "APPLICATIONFACTFINDNUMBER")
    strInQuotationNumber = objXmlAssist.GetTagValue(xmlElement, "QUOTATIONNUMBER")
    strCountryOfOrigin = objXmlAssist.GetTagValue(xmlElement, "COUNTRYOFORIGIN")
    strLanguage = objXmlAssist.GetTagValue(xmlElement, "LANGUAGE")

    If Len(strInApplicationNumber) = 0 Or _
        Len(strInApplicationFactFindNumber) = 0 Or _
        Len(strInQuotationNumber) = 0 Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingPrimaryTag, _
        "APPLICATIONNUMBER =" & strInApplicationNumber & _
        "; APPLICATIONFACTFINDNUMBER =" & strInApplicationFactFindNumber & _
        "; QUOTATIONNUMBER =" & strInQuotationNumber & _
        "; COUNTRYOFORIGIN =" & strCountryOfOrigin & _
        "; LANGUAGE =" & strLanguage

    End If
    
    Set objSQLAssist = New SQLAssist
    
    Dim strSQL As String
                   
    strSQL = "SELECT * FROM GETQUOTATIONSUMMARY WHERE APPLICATIONNUMBER = "
    strSQL = strSQL & objSQLAssist.FormatString(strInApplicationNumber)
    strSQL = strSQL & " AND APPLICATIONFACTFINDNUMBER = " & strInApplicationFactFindNumber
    strSQL = strSQL & " AND QUOTATIONNUMBER = " & strInQuotationNumber
    strSQL = strSQL & " AND COUNTRYOFORIGIN = " & strCountryOfOrigin
    strSQL = strSQL & " AND LANGUAGE = " & strLanguage
    strSQL = strSQL & " ORDER BY LOANCOMPONENTSEQUENCENUMBER"
            
    'IVW - Disconnecting record sets
    
    Dim cmd As ADODB.Command
    Dim conn As ADODB.Connection
    Dim rstThisRecSet As ADODB.Recordset
    
    Set cmd = New ADODB.Command
    Set conn = New ADODB.Connection
    Set rstThisRecSet = New ADODB.Recordset

    cmd.CommandText = strSQL
    
    rstThisRecSet.CursorLocation = adUseClient
    rstThisRecSet.CursorType = adOpenForwardOnly
    rstThisRecSet.LockType = adLockReadOnly
    
    conn.Open objIADOAssist.GetConnStr
    cmd.ActiveConnection = conn
            
    Set rstThisRecSet.Source = cmd
    
    rstThisRecSet.Open
    
    Set rstThisRecSet.ActiveConnection = Nothing
    
 
    'IVW
    
    Set objIomCMClassDef = New omCMClassDef
    Set objIomMPClassDef = New omMPClassDef
    
    Set xmlQuotationSummaryDefDoc = objIomCMClassDef.LoadQuotationSummaryViewData()
    Set xmlLifeBenefitDefDoc = objIomCMClassDef.LoadLifeBenefitData(QuotationSummaryView)
    Set xmlLoanComponentDefDoc = objIomCMClassDef.LoadLoanComponentData(QuotationSummaryView)
    Set xmlMortProdLangDefDoc = objIomMPClassDef.LoadQuotationSummaryViewData()
                
    Set xmlMortProdLangTableElem = xmlMortProdLangDefDoc.getElementsByTagName("TABLENAME").Item(0)
    
    If xmlMortProdLangTableElem Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeMissingPrimaryTag, _
                                "MORTGAGEPRODUCTLANGUAGE class definition not found"
    End If
    
    Set xmlLoanCompTableElem = xmlLoanComponentDefDoc.getElementsByTagName("TABLENAME").Item(0)
    
    If xmlLoanCompTableElem Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeMissingPrimaryTag, _
                                "LOANCOMPONENT class definition not found"
    End If
        
    For Each xmlNode In xmlMortProdLangTableElem.childNodes
        If xmlNode.nodeType = NODE_ELEMENT Then
            xmlLoanCompTableElem.appendChild xmlNode
        End If
    Next
    
    Dim intLoanCompNo As Integer
    Dim intCurrentLoanCompNo As Integer
    Dim intCurrentQuoteNo As Integer
    Dim intQuoteNo As Integer
    Dim intFirstLoanCompNo As Integer
    
    If rstThisRecSet.EOF Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotfound
    End If

    ' AD 05/10/01 Removed MoveFirst as it's a ForwardOnly cursor. SYS2718.
    
    Set fldFields = rstThisRecSet.Fields
    
    Do While Not rstThisRecSet.EOF
    
        intCurrentQuoteNo = fldFields("QUOTATIONNUMBER")
    
        If intCurrentQuoteNo <> intQuoteNo Then
            
            intQuoteNo = intCurrentQuoteNo
            intFirstLoanCompNo = fldFields("LOANCOMPONENTSEQUENCENUMBER").Value
            
            Set xmlQuoteSummaryElem = objIDOAssist.GetXMLFromRecordSet(rstThisRecSet, xmlQuotationSummaryDefDoc)
            xmlOut.appendChild xmlQuoteSummaryElem
            
            Set xmlLoanCompListElem = xmlOut.createElement("LOANCOMPONENTLIST")
            xmlQuoteSummaryElem.appendChild xmlLoanCompListElem
            Set xmlLifeBenefitListElem = xmlOut.createElement("LIFEBENEFITLIST")
            xmlQuoteSummaryElem.appendChild xmlLifeBenefitListElem
        End If
        
        ' Get the loan component number and if different add the loan component details to the
        ' loan component list
        intCurrentLoanCompNo = fldFields("LOANCOMPONENTSEQUENCENUMBER").Value
        
        If intCurrentLoanCompNo <> intLoanCompNo Then
            intLoanCompNo = intCurrentLoanCompNo
            objIDOAssist.GetXMLFromRecordSet rstThisRecSet, _
                                             xmlLoanComponentDefDoc, _
                                             xmlLoanCompListElem
            
        End If
        
        ' As the benefit data is repeated for each loan component only extract it for the
        ' first loan component
        ' AY 21/03/00 - don't generate XML if there are no life subquotes
        ' START: MAR632 - Maha T
            'If intFirstLoanCompNo = intLoanCompNo Then
            '    If Not IsNull(fldFields("LIFESUBQUOTENUMBER").Value) Then
            '       objIDOAssist.GetXMLFromRecordSet rstThisRecSet, _
            '                                 xmlLifeBenefitDefDoc, _
            '                                 xmlLifeBenefitListElem
            '    End If
            'End If
        ' END: MAR632
        rstThisRecSet.MoveNext
    Loop
        
    Set IQuotationDO_GetQuotationSummary = xmlQuoteSummaryElem
    
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Set cmd = Nothing
    rstThisRecSet.Close
    Set rstThisRecSet = Nothing
    conn.Close
    Set conn = Nothing
    Set objContext = Nothing

    Exit Function

GetQuotationSummaryVbErr:
        
    ' Close the recordset if still open
    If Not rstThisRecSet Is Nothing Then
        If rstThisRecSet.State = adStateOpen Then
            rstThisRecSet.Close
        End If
    End If

    ' Close the connection if still open
    If Not conn Is Nothing Then      'JLD SYS3499
        If conn.State = adStateOpen Then
            conn.Close
        End If
    End If
    
    Set cmd = Nothing
    Set conn = Nothing
    Set rstThisRecSet = Nothing

    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If
    
    If Not objContext Is Nothing Then
        If objErrAssist.IsSystemError Then
            objContext.SetAbort
        Else
            objContext.SetComplete
        End If
    End If
    
    Set objErrAssist = Nothing
    Set objContext = Nothing
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.number, Err.Source, Err.Description

End Function

Private Function IQuotationDO_StoredQuoteFindList(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description: Counts the number of quotations a mortgage sub quote is attached to
' pass:
'   vxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'
'
'   Input XML
'
'   <QUOTATION>
'       <APPLICATIONNUMBER>Value</APPLICATIONNUMBER>
'       <APPLICATIONFACTFINDNUMBER>Value</APPLICATIONFACTFINDNUMBER>
'        <MORTGAGESUBQUOTENUMBER>Value</MORTGAGESUBQUOTENUMBER>
'   </QUOTATION>
'------------------------------------------------------------------------------------------
On Error GoTo StoredQuoteFindListVbErr

    Const strFunctionName As String = "IQuotationDO_StoredQuoteFindList"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
        
    Dim objErrAssist As New ErrAssist
    Dim objIDOAssist As IDOAssist
    Dim objIADOAssist As IADOAssist
    Dim objXmlAssist As New XMLAssist
    Dim objSQLAssist As New SQLAssist
    Dim xmlLoanCompDefDoc As FreeThreadedDOMDocument40
    Dim xmlStoredQuoteDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IomCMClassDef
    Dim xmlListElem As IXMLDOMElement
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlElement As IXMLDOMElement
    Dim xmlStoredQuoteNode As IXMLDOMNode
    Dim xmlLoanCompListElem As IXMLDOMElement
                 
    If objContext Is Nothing Then
        Set objIDOAssist = New DOAssist
        Set objIADOAssist = New ADOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
        Set objIADOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    End If
     
    Dim strApplicationNo As String
    Dim strApplicationFFNo As String
    
    strApplicationNo = objXmlAssist.GetTagValue(vxmlTableElement, "APPLICATIONNUMBER")
    strApplicationFFNo = objXmlAssist.GetTagValue(vxmlTableElement, _
                                                  "APPLICATIONFACTFINDNUMBER")
                                                     
    If Len(strApplicationNo) = 0 Or Len(strApplicationFFNo) = 0 Then
        Dim strDetails As String
        
        strDetails = "APPLICATIONNUMBER = " & strApplicationNo
        strDetails = strDetails & "; APPLICATIONFACTFINDNUMBER = " & strApplicationFFNo
        
        objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeMissingPrimaryTag, _
                                strDetails
    End If
    
    Dim strSQL As String
    
    strSQL = "SELECT * FROM STOREDQUOTEFINDLIST WHERE APPLICATIONNUMBER = "
    strSQL = strSQL & objSQLAssist.FormatString(strApplicationNo)
    strSQL = strSQL & " AND APPLICATIONFACTFINDNUMBER = " & strApplicationFFNo
    strSQL = strSQL & " ORDER BY QUOTATIONNUMBER, LOANCOMPONENTSEQUENCENUMBER"
    
    'IVW - Disconnecting record sets
    
    Dim cmd As ADODB.Command
    Dim conn As ADODB.Connection
    Dim rstThisRecSet As ADODB.Recordset
    
    Set cmd = New ADODB.Command
    Set conn = New ADODB.Connection
    Set rstThisRecSet = New ADODB.Recordset

    cmd.CommandText = strSQL
    
    rstThisRecSet.CursorLocation = adUseClient
    rstThisRecSet.CursorType = adOpenForwardOnly
    rstThisRecSet.LockType = adLockReadOnly
    
    conn.Open objIADOAssist.GetConnStr
    cmd.ActiveConnection = conn
            
    Set rstThisRecSet.Source = cmd
    
    rstThisRecSet.Open
    
    Set rstThisRecSet.ActiveConnection = Nothing
    
 
    'IVW
    
    Set objIClassDef = New omCMClassDef
    
    Set xmlLoanCompDefDoc = objIClassDef.LoadLoanComponentData(LoanComponentStoredQuoteView)
    Set xmlStoredQuoteDoc = objIClassDef.LoadStoredQuoteFindListViewData()
    
    Set xmlListElem = xmlOut.createElement("STOREDQUOTATIONLIST")
    xmlOut.appendChild xmlListElem
      
    If rstThisRecSet.EOF Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotfound
    End If
    
    Dim intQuotationNo As Integer
    Dim intCurrentQuotationNo As Integer
    Dim intLoanCompSeqNo As Integer

    ' AD 05/10/01 Removed MoveFirst as it's a ForwardOnly cursor. SYS2718.
    
    Do While Not rstThisRecSet.EOF
        intCurrentQuotationNo = rstThisRecSet("QUOTATIONNUMBER").Value
        
        If intCurrentQuotationNo <> intQuotationNo Then
            Set xmlStoredQuoteNode = objIDOAssist.GetXMLFromRecordSet(rstThisRecSet, _
                                                                      xmlStoredQuoteDoc, _
                                                                      xmlListElem)
            intQuotationNo = intCurrentQuotationNo
            Set xmlLoanCompListElem = Nothing
        End If
        
        ' Add in the loan component data for this quote
        If intCurrentQuotationNo = intQuotationNo Then
                
            ' Check if we have a loan component
            If Not IsNull(rstThisRecSet("LOANCOMPONENTSEQUENCENUMBER").Value) Then
                
                ' If it is the first one create the list element to attach to
                If xmlLoanCompListElem Is Nothing Then
                    Set xmlLoanCompListElem = xmlOut.createElement("LOANCOMPONENTLIST")
                    xmlStoredQuoteNode.appendChild xmlLoanCompListElem
                End If
                                
                objIDOAssist.GetXMLFromRecordSet rstThisRecSet, _
                                                 xmlLoanCompDefDoc, _
                                                 xmlLoanCompListElem
            End If
        End If
        rstThisRecSet.MoveNext
    Loop
    
    Set IQuotationDO_StoredQuoteFindList = xmlOut.documentElement
    
    Set objErrAssist = Nothing
    Set objIDOAssist = Nothing
    Set objIADOAssist = Nothing
    Set objXmlAssist = Nothing
    Set objSQLAssist = Nothing
    Set xmlLoanCompDefDoc = Nothing
    Set xmlStoredQuoteDoc = Nothing
    Set objIClassDef = Nothing
    Set xmlListElem = Nothing
    Set xmlOut = Nothing
    Set xmlElement = Nothing
    Set xmlStoredQuoteNode = Nothing
    Set cmd = Nothing
        rstThisRecSet.Close
        Set rstThisRecSet = Nothing
        conn.Close
    Set conn = Nothing
    
    Set xmlLoanCompListElem = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

StoredQuoteFindListVbErr:
    
    ' Close the recordset if still open
        If Not rstThisRecSet Is Nothing Then
                If rstThisRecSet.State = adStateOpen Then
                        rstThisRecSet.Close
                End If
        End If


    ' Close the connection if still open
    If conn.State = adStateOpen Then
        conn.Close
    End If
    
    Set cmd = Nothing
    Set conn = Nothing
    Set rstThisRecSet = Nothing
    
    Set objIDOAssist = Nothing
    Set objIADOAssist = Nothing
    Set objXmlAssist = Nothing
    Set objSQLAssist = Nothing
    Set xmlLoanCompDefDoc = Nothing
    Set xmlStoredQuoteDoc = Nothing
    Set objIClassDef = Nothing
    Set xmlListElem = Nothing
    Set xmlOut = Nothing
    Set xmlElement = Nothing
    Set xmlStoredQuoteNode = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
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
    Err.Raise Err.number, Err.Source, Err.Description

End Function

Private Sub IQuotationDO_Update(ByVal vxmlTableElement As IXMLDOMElement)

' header ----------------------------------------------------------------------------------
' description:
'   Update a single instance of the persistant data associated with this
'   data object
' pass:
'   vxmlTableElement     xml element containing the request
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo UpdateVbErr

    Const strFunctionName As String = "IQuotationDO_Update"
        
    Dim objErrAssist As New ErrAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIomCMClassDef As IomCMClassDef
    Dim objIDOAssist As IDOAssist
    
    If objContext Is Nothing Then
      Set objIDOAssist = New DOAssist
    Else
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set objIomCMClassDef = New omCMClassDef
    Set xmlClassDefDoc = objIomCMClassDef.LoadQuotationData()
    objIDOAssist.Update vxmlTableElement, xmlClassDefDoc
    
    Set objIomCMClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    Set objErrAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Sub

UpdateVbErr:
        
    Set objIomCMClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIDOAssist = Nothing
    
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
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
    Err.Raise Err.number, Err.Source, Err.Description

End Sub

Private Sub IQuotationDO_AddDerivedData(ByVal vxmlData As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   XML elements must be created for any derived values as specified.
'   Add any derived values to XML. E.g. data type 'double' fields will
'   need to be formatted as strings to required precision & rounding.
' pass:
'   vxmlData          base XML node
'                       as:
'                           <tablename>
'                               <element1>element1 value</element1>
'                               <elementn>elementn value</elementn>
' return:                n/a
'------------------------------------------------------------------------------------------
On Error GoTo AddDerivedDataVBErr

    Const strFunctionName As String = "IQuotationDO_AddDerivedData"

    Dim objErrAssist As New ErrAssist

    Set objErrAssist = Nothing
    
    Exit Sub
    
AddDerivedDataVBErr:
    
    Set objErrAssist = Nothing
    
    Err.Raise Err.number, Err.Source, Err.Description

End Sub

Private Function IQuotationDO_GetAcceptedBCSubQuoteNumber(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   this data object
' pass:
'   vxmlRequest     xml Request data stream containing data to which identifies
'                   the instance of the persistant data to be retrieved
'   vxmlClassDef    xml Class Def to parse vxmlRequest against
' return:
'   GetDataEx       string containing XML data stream representation of
'                   data retrieved
' Raise Errors: if record not found, raise omiga4RecordNotFound
'------------------------------------------------------------------------------------------
On Error GoTo GetAcceptedBCSubQuoteNumberVbErr

    Const strFunctionName As String = "IQuotationDO_GetAcceptedBCSubQuoteNumber"

    Dim objErrAssist As ErrAssist
    Set objErrAssist = New ErrAssist
        
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
        
    Dim objIADOAssist As IADOAssist
    Dim objIDOAssist As IDOAssist
    Dim objSQLAssist As SQLAssist
    Set objSQLAssist = New SQLAssist
    Dim xmlElement As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    Dim objXmlAssist As XMLAssist
    Set objXmlAssist = New XMLAssist
    Dim xmlAcceptedBCSubQuoteNumberDoc As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlAcceptedBCSubQuoteNumberElem As IXMLDOMElement

    Dim objIomCMClassDef As IomCMClassDef
    
    If objContext Is Nothing Then
        Set objIADOAssist = New ADOAssist
        Set objIDOAssist = New DOAssist
    Else
        Set objIADOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
                
    
    If vXMLRequest.nodeName = "QUOTATION" Then
       Set xmlElement = vXMLRequest
    Else
        Set xmlElement = vXMLRequest.getElementsByTagName("QUOTATION").Item(0)
    End If
    
    If xmlElement Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "QUOTATION tag not found"
    End If
    
    Dim strInApplicationNumber As String
    Dim strInApplicationFactFindNumber As String

    'Need to get the values for these fields out of the XML passed in
    strInApplicationNumber = objXmlAssist.GetTagValue(xmlElement, "APPLICATIONNUMBER")
    strInApplicationFactFindNumber = objXmlAssist.GetTagValue(xmlElement, "APPLICATIONFACTFINDNUMBER")

    If Len(strInApplicationNumber) = 0 Or _
        Len(strInApplicationFactFindNumber) = 0 Then

        objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingPrimaryTag, _
        "APPLICATIONNUMBER =" & strInApplicationNumber & _
        "; APPLICATIONFACTFINDNUMBER =" & strInApplicationFactFindNumber

    End If
                                
    Dim strSQL As String
                   
    strSQL = "SELECT QUOTATION.BCSUBQUOTENUMBER FROM APPLICATIONFACTFIND,QUOTATION "
    strSQL = strSQL & " WHERE APPLICATIONFACTFIND.APPLICATIONNUMBER = QUOTATION.APPLICATIONNUMBER"
    strSQL = strSQL & " AND APPLICATIONFACTFIND.APPLICATIONFACTFINDNUMBER = QUOTATION.APPLICATIONFACTFINDNUMBER "
    strSQL = strSQL & " AND APPLICATIONFACTFIND.ACCEPTEDQUOTENUMBER = QUOTATION.QUOTATIONNUMBER "
    strSQL = strSQL & " AND APPLICATIONFACTFIND.APPLICATIONNUMBER = " & objSQLAssist.FormatString(strInApplicationNumber)
    strSQL = strSQL & " AND APPLICATIONFACTFIND.APPLICATIONFACTFINDNUMBER = " & strInApplicationFactFindNumber
    
    'IVW - Disconnecting record sets
    Dim cmd As ADODB.Command
    Dim conn As ADODB.Connection
    Dim rstThisRecSet As ADODB.Recordset
    
    Set cmd = New ADODB.Command
    Set conn = New ADODB.Connection
    Set rstThisRecSet = New ADODB.Recordset

    cmd.CommandText = strSQL
    
    rstThisRecSet.CursorLocation = adUseClient
    rstThisRecSet.CursorType = adOpenForwardOnly
    rstThisRecSet.LockType = adLockReadOnly
    
    conn.Open objIADOAssist.GetConnStr
    cmd.ActiveConnection = conn
            
    Set rstThisRecSet.Source = cmd
    
    rstThisRecSet.Open
    
    Set rstThisRecSet.ActiveConnection = Nothing
    
 
    'IVW
    
    Set objIomCMClassDef = New omCMClassDef

    Set xmlAcceptedBCSubQuoteNumberDoc = objIomCMClassDef.LoadQuotationData(BCSubQuoteNumber)
    
    If Not rstThisRecSet.EOF Then
        Set xmlAcceptedBCSubQuoteNumberElem = objIDOAssist.GetXMLFromRecordSet(rstThisRecSet, xmlAcceptedBCSubQuoteNumberDoc)
    Else
        objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotfound
    End If
        
    Set IQuotationDO_GetAcceptedBCSubQuoteNumber = xmlAcceptedBCSubQuoteNumberElem
        
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Set cmd = Nothing
        rstThisRecSet.Close
        Set rstThisRecSet = Nothing
        conn.Close
    Set conn = Nothing
    

    Exit Function

GetAcceptedBCSubQuoteNumberVbErr:
        
    objErrAssist.AddToErrSource (strFunctionName)
    
    If objErrAssist.IsSystemError = True Then
        objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If

        ' Close the recordset if still open
        If Not rstThisRecSet Is Nothing Then
                If rstThisRecSet.State = adStateOpen Then
                        rstThisRecSet.Close
                End If
        End If


    ' Close the connection if still open
    If conn.State = adStateOpen Then
        conn.Close
    End If
    
    Set cmd = Nothing
    Set conn = Nothing
    Set rstThisRecSet = Nothing
    
    If Not objContext Is Nothing Then
        If objErrAssist.IsSystemError Then
            objContext.SetAbort
        Else
            objContext.SetComplete
        End If
    End If
    
    Set objErrAssist = Nothing
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.number, Err.Source, Err.Description

End Function

Private Function IQuotationDO_GetQuotationLoanSummary(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for all instances of the persistant data associated with
'   this data object for the values supplied
' pass:
'   vxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo GetQuotationLoanSummaryVbErr

    Const strFunctionName As String = "IQuotationDO_GetQuotationLoanSummary"
    
        
    Dim xmlClassDefDoc As FreeThreadedDOMDocument40
    Dim objIClassDef As IomCMClassDef
    Dim objIADOAssist As IADOAssist
    Dim objXmlAssist As XMLAssist
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlDataNode As IXMLDOMElement
    Dim xmlCostType As IXMLDOMElement
    Dim xmlAmount As IXMLDOMElement
    Dim xmlElement As IXMLDOMElement
    
    Dim adoConnection As ADODB.Connection
    Dim adoCommand As ADODB.Command
    Dim adoRecordset As ADODB.Recordset
    Dim adoParameter As ADODB.Parameter
    
    Dim strApplicationNumber As String
    Dim strApplicationFFNumber As String
    Dim strQuotationNo As String
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
                                                       
    strQuotationNo = objXmlAssist.GetNodeValue(vxmlTableElement, _
                                               ".//QUOTATIONNUMBER")
                                                       
    
    Set objIADOAssist = New ADOAssist
    
    Set adoConnection = New ADODB.Connection
    Set adoCommand = New ADODB.Command
    Set adoRecordset = New ADODB.Recordset
    
    'BM0464  Set up the stored procedure (which uses view GETQUOTATIONLOANSUMMARY)
    adoCommand.CommandType = adCmdStoredProc
    adoCommand.CommandText = "USP_GETQUOTATIONLOANSUMMARYVIEW"
    
    ' Set up the parameters for the stored procedure.
    Set adoParameter = adoCommand.CreateParameter("@sApplicationNumber", adBSTR, adParamInput, Len(strApplicationNumber), strApplicationNumber)
    adoCommand.Parameters.Append adoParameter

    Set adoParameter = adoCommand.CreateParameter("@sApplicationFactFindNumber", adBSTR, adParamInput, Len(strApplicationFFNumber), strApplicationFFNumber)
    adoCommand.Parameters.Append adoParameter
    
    Set adoParameter = adoCommand.CreateParameter("@sQuotationNumber", adBSTR, adParamInput, Len(strQuotationNo), strQuotationNo)
    adoCommand.Parameters.Append adoParameter
    
    ' Set Recordset properties. Must be adUseClient for disconnected recordset
    adoRecordset.CursorLocation = adUseClient
    adoRecordset.CursorType = adOpenForwardOnly
    adoRecordset.LockType = adLockReadOnly
    
    adoConnection.CursorLocation = adUseClient  ' BM0454 Disconnected recordset

    intRetries = objIADOAssist.GetRetries()
    intMaxAttempts = 1 + intRetries
    
    blnOpenedOk = False
    intAttempt = 1
    strConnString = objIADOAssist.GetConnStr()
    
    ' Keep trying to open the connection for the specified number of retries
    While (blnOpenedOk = False) And (intAttempt <= intMaxAttempts)
    
        adoConnection.Errors.Clear
        
        adoConnection.Open strConnString
        
        If adoConnection.Errors.count = 0 Then
            blnOpenedOk = True
        End If
 
        intAttempt = intAttempt + 1
    
    Wend
    
    If blnOpenedOk = False Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeUnableToConnect
    End If
    
    adoCommand.ActiveConnection = adoConnection
    
    'BM0464 Execute the stored procedure to populate the record set
    Set adoRecordset = adoCommand.Execute
    
    'Disconnect the recordset
    Set adoRecordset.ActiveConnection = Nothing
    adoConnection.Close
    
    If adoRecordset.EOF Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotfound
    End If
     
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    
    Set objIClassDef = New omCMClassDef
    Set xmlClassDefDoc = objIClassDef.LoadQuotationLoanSummaryViewData()
    
    ' AD 05/10/01 Removed MoveFirst as it's a ForwardOnly cursor. SYS2718.
    
    ' Convert first record to xml
    Set xmlDataNode = m_objIDOAssist.GetXMLFromRecordSet(adoRecordset, xmlClassDefDoc)
    
    ' Remove the ONEOFFCOSTTYPE element and convert the AMOUNT element to INDEMNITYPREMIUM
    Set xmlCostType = m_objXmlAssist.GetNode(xmlDataNode, ".//ONEOFFCOSTTYPE")
    Set xmlAmount = m_objXmlAssist.GetNode(xmlDataNode, ".//AMOUNT")
    
    xmlDataNode.removeChild xmlCostType
    xmlDataNode.removeChild xmlAmount
    
    Set xmlElement = xmlOut.createElement("INDEMNITYPREMIUM")
    
    ' First record will be MIG or nothing
    If Len(xmlCostType.Text) > 0 Then
        xmlElement.Text = xmlAmount.Text
    Else
        xmlElement.Text = xmlAmount.Text
    End If
    
    xmlDataNode.appendChild xmlElement
    
    ' Move to the next record to get the Indemnity Amount
    adoRecordset.MoveNext
    
    Set xmlElement = xmlOut.createElement("INDEMNITYAMOUNT")

    ' If not there send back zero
    If adoRecordset.EOF = False Then
        xmlElement.Text = adoRecordset.Fields("AMOUNT").Value
    Else
        xmlElement.Text = 0
    End If
            
    xmlDataNode.appendChild xmlElement

            
    adoRecordset.Close
    
    Set IQuotationDO_GetQuotationLoanSummary = xmlDataNode
    
    Set objIClassDef = Nothing
    Set xmlClassDefDoc = Nothing
    Set objIADOAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlDataNode = Nothing
    Set adoConnection = Nothing
    Set adoCommand = Nothing
    Set adoRecordset = Nothing
    Set adoParameter = Nothing
    Set xmlOut = Nothing
    
    m_objContext.SetComplete
    
    
    Exit Function

GetQuotationLoanSummaryVbErr:
    
    
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
    Set objIADOAssist = Nothing
    Set objXmlAssist = Nothing
    Set adoConnection = Nothing
    Set adoCommand = Nothing
    Set adoRecordset = Nothing
    Set adoParameter = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing

    m_objErrAssist.AddToErrSource (strFunctionName)
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.number, Err.Description
    End If
    
    If m_objErrAssist.IsSystemError Then
        m_objContext.SetAbort
    Else
        m_objContext.SetComplete
    End If
    
    '   re-raise error for business object to interpret as appropriate
    Err.Raise Err.number, Err.Source, Err.Description

End Function

Private Sub ObjectControl_Activate()

    Set m_objContext = GetObjectContext()
    Set m_objErrAssist = New ErrAssist
    Set m_objXmlAssist = New XMLAssist
    Set m_objIDOAssist = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    Set m_objSQLAssist = New SQLAssist
End Sub

Private Function ObjectControl_CanBePooled() As Boolean

    ObjectControl_CanBePooled = False

End Function

Private Sub ObjectControl_Deactivate()

    Set m_objContext = Nothing
    Set m_objErrAssist = Nothing
    Set m_objXmlAssist = Nothing
    Set m_objIDOAssist = Nothing
    Set m_objSQLAssist = Nothing

End Sub

'BMIDS624 GHun 22/10/2003
Private Function IQuotationDO_HaveRatesChanged(ByVal vXMLRequest As IXMLDOMElement) As IXMLDOMNode
    Const cstrFunctionName  As String = "HaveRatesChanged"
    
    Dim sApplicationNumber          As String
    Dim sApplicationFactFindNumber  As String
    Dim sQuotationNumber            As String               ' MAR1500
    Dim adoConnection               As ADODB.Connection
    Dim adoCommand                  As ADODB.Command
    Dim adoReturnParam              As ADODB.Parameter
    Dim xmlQuote                    As IXMLDOMElement
    Dim xmlDoc                      As FreeThreadedDOMDocument40

On Error GoTo HaveRatesChangedExit
    
    sApplicationNumber = xmlGetMandatoryNodeText(vXMLRequest, "//APPLICATIONNUMBER")
    sApplicationFactFindNumber = xmlGetMandatoryNodeText(vXMLRequest, "//APPLICATIONFACTFINDNUMBER")
    
    'MAR1500 Use Quotation number if it has been passed in.
    '        Otherwise Active/Accepted quote will be used in stored procedure.
    sQuotationNumber = xmlGetNodeText(vXMLRequest, "//QUOTATIONNUMBER")

    Set adoConnection = New ADODB.Connection
    With adoConnection
        .ConnectionString = adoGetDbConnectString
        .Open
    End With
                      
    'Create the parameters and call the stored proc
    Set adoCommand = New ADODB.Command
    With adoCommand
        .ActiveConnection = adoConnection
        .CommandType = adCmdStoredProc
        .CommandText = "USP_HAVERATESCHANGED"
        Set adoReturnParam = .CreateParameter("@ReturnValue", adInteger, adParamReturnValue)
        .Parameters.Append adoReturnParam
        .Parameters.Append .CreateParameter("@p_ApplicationNumber", adBSTR, adParamInput, 12, sApplicationNumber)
        .Parameters.Append .CreateParameter("@p_ApplicationFactFindNumber", adInteger, adParamInput, , sApplicationFactFindNumber)
        
        'MAR1500
        If (sQuotationNumber <> "") Then
            .Parameters.Append .CreateParameter("@p_QuotationNumber", adInteger, adParamInput, , sQuotationNumber)
        End If
                
        .Execute , , adExecuteNoRecords
    End With
    
    'Close the database connection
    adoConnection.Close
    Set adoCommand.ActiveConnection = Nothing
    
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    
    'The return value holds the RatesChanged indicator
    Set xmlQuote = xmlDoc.createElement("QUOTATION")
    If adoReturnParam.Value = 1 Then
        xmlQuote.setAttribute "RATESINCONSISTENT", 1
    Else
        xmlQuote.setAttribute "RATESINCONSISTENT", 0
    End If

    Set IQuotationDO_HaveRatesChanged = xmlQuote

HaveRatesChangedExit:
    Set xmlDoc = Nothing
    Set adoCommand = Nothing
    Set adoConnection = Nothing
    Set adoReturnParam = Nothing
    Set xmlQuote = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)
End Function
'BMIDS624 End

