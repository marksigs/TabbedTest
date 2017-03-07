VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "PaymentProtectionSubQuoteDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'
'Workfile:      PaymentProtectionSubQuoteDO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling
'
'Description:   Payment Protection Sub Quote Data Object
'
'Dependencies:
'
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date     Description
'PSC    02/11/99 Created
'PSC    24/01/00 Amended FindList to return the PaymentProtectionDetails record as well
'                and to use IDOAssist interface
'IVW    09/11/00 Amended to disconnect record sets more efficiently
'AD     05/10/01 Removed the MoveFirst's as it's a ForwardOnly cursor. SYS2718.
'------------------------------------------------------------------------------------------

Option Explicit
 
 '=============================================
 'Variable Declaration Section
 '=============================================
Private Const cstrSUBQUOTE_NAME = "PAYMENTPROTECTIONSUBQUOTE"
 
 '=============================================
 'Constant Declaration Section
 '=============================================

Public Function CreateSubQuoteDetails(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description: Creates the Payment Protection Sub Quote Details consisting of the
'              Payment Protection Sub quote and the Payment Protection Details
' pass:
'   vstrXMLRequest  xml Request data stream containing data to be persisted
'------------------------------------------------------------------------------------------
On Error GoTo CreateSubQuoteDetailsVbErr

    Const strFunctionName As String = "CreateSubQuoteDetails"
   
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    Dim objDOAssist As DOAssist
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlTableElem As IXMLDOMElement
    Dim xmlSubQuoteNumElem As IXMLDOMElement
    Dim xmlElement As IXMLDOMElement
    Dim xmlSubQuoteElem As IXMLDOMElement
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Set xmlTableElem = xmlOut.createElement(cstrSUBQUOTE_NAME)
    xmlOut.appendChild xmlTableElem
               
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    Set xmlSubQuoteElem = xmlIn.getElementsByTagName(cstrSUBQUOTE_NAME).Item(0)
    
    If xmlSubQuoteElem Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeMissingPrimaryTag, _
                                cstrSUBQUOTE_NAME & " tag not found"
    End If
    
    
    Dim strApplicationNo As String
    Dim strApplicationFFNo As String
    Dim strTagName As String
    Dim lngPPSequenceNo As Long

    strTagName = "APPLICATIONNUMBER"
    strApplicationNo = objXmlAssist.GetTagValue(xmlSubQuoteElem, strTagName)
    
    strTagName = "APPLICATIONFACTFINDNUMBER"
    strApplicationFFNo = objXmlAssist.GetTagValue(xmlSubQuoteElem, strTagName)
    
    lngPPSequenceNo = GetHighestSubQuoteNumber(strApplicationNo, _
                                               strApplicationFFNo)
                                               
    lngPPSequenceNo = lngPPSequenceNo + 1
    
    ' Add the sub quote number into payment protection sub quote
    strTagName = "PPSUBQUOTENUMBER"
    Set xmlElement = xmlSubQuoteElem.getElementsByTagName(strTagName).Item(0)
    
    If xmlElement Is Nothing Then
        Set xmlElement = xmlIn.createElement(strTagName)
        xmlSubQuoteElem.appendChild xmlElement
    End If

    If Len(xmlElement.Text) = 0 Then
        xmlElement.Text = CStr(lngPPSequenceNo)
    End If
    
    ' Copy the element into the output xml
    xmlTableElem.appendChild xmlElement.cloneNode(True)
        
    If objContext Is Nothing Then
        Set objDOAssist = New DOAssist
    Else
        Set objDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    CreatePaymentProtectionSubQuote xmlSubQuoteElem.xml
    objXmlAssist.ChangeNodeName xmlSubQuoteElem, cstrSUBQUOTE_NAME, "PAYMENTPROTECTIONDETAILS"
    CreatePaymentProtectionDetails xmlSubQuoteElem.xml
    
    CreateSubQuoteDetails = xmlOut.xml
    
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objDOAssist = Nothing
    Set xmlIn = Nothing
    Set xmlElement = Nothing
    Set xmlSubQuoteElem = Nothing
    Set xmlOut = Nothing
    Set xmlTableElem = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

CreateSubQuoteDetailsVbErr:

    Set objXmlAssist = Nothing
    Set objDOAssist = Nothing
    Set xmlIn = Nothing
    Set xmlElement = Nothing
    Set xmlSubQuoteElem = Nothing
    Set xmlOut = Nothing
    Set xmlTableElem = Nothing
    
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

Public Function GetSubQuoteDetails(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description: Gets the Payment Protection Sub Quote Details consisting of the
'              Payment Protection Sub quote and the Payment Protection Details
' pass:
'   vstrXMLRequest  xml Request data stream containing data to which identifies
'                   the instance of the persistant data to be retrieved
' return:           string containing XML data stream representation of
'                   data retrieved
'------------------------------------------------------------------------------------------
On Error GoTo GetSubQuoteDetailsVbErr

    Const strFunctionName As String = "GetSubQuoteDetails"
        
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim objSQLAssist As New SQLAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim strSQL As String
    Dim strXML As String

    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlDoc As New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    Dim xmlIn As FreeThreadedDOMDocument40
    
    Dim xmlNode As IXMLDOMNode
    Dim xmlSubQuoteElem As IXMLDOMElement
    
    Dim objADOAssist As ADOAssist
    Dim objDOAssist As DOAssist
    
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    Dim strApplicationNumber As String
    Dim strApplicationFFNumber As String
    Dim strPPSubQuoteNumber As String
    
    strApplicationNumber = _
        objXmlAssist.GetTagValue(xmlIn.documentElement, "APPLICATIONNUMBER")
    strApplicationFFNumber = _
        objXmlAssist.GetTagValue(xmlIn.documentElement, "APPLICATIONFACTFINDNUMBER")
    strPPSubQuoteNumber = _
        objXmlAssist.GetTagValue(xmlIn.documentElement, "PPSUBQUOTENUMBER")
    
    If Len(strApplicationNumber) = 0 Or Len(strApplicationFFNumber) = 0 Or _
       Len(strPPSubQuoteNumber) = 0 Then
        objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeMissingPrimaryTag, _
                                "APPLICATIONNUMBER = " & strApplicationNumber & _
                                "; APPLICATIONFACTFINDNUMBER = " & strApplicationFFNumber & _
                                "; PPSUBQUOTENUMBER = " & strPPSubQuoteNumber
    End If
    
    If objContext Is Nothing Then
        Set objADOAssist = New ADOAssist
        Set objDOAssist = New DOAssist
    Else
        Set objADOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
        Set objDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    strSQL = "SELECT * FROM PAYMENTPROTECTIONSUBQUOTE, PAYMENTPROTECTIONDETAILS WHERE"
    strSQL = strSQL & " PAYMENTPROTECTIONSUBQUOTE.APPLICATIONNUMBER = "
    strSQL = strSQL & " PAYMENTPROTECTIONDETAILS.APPLICATIONNUMBER AND"
    strSQL = strSQL & " PAYMENTPROTECTIONSUBQUOTE.APPLICATIONFACTFINDNUMBER = "
    strSQL = strSQL & " PAYMENTPROTECTIONDETAILS.APPLICATIONFACTFINDNUMBER AND"
    strSQL = strSQL & " PAYMENTPROTECTIONSUBQUOTE.PPSUBQUOTENUMBER = "
    strSQL = strSQL & " PAYMENTPROTECTIONDETAILS.PPSUBQUOTENUMBER AND"
    strSQL = strSQL & " PAYMENTPROTECTIONSUBQUOTE.APPLICATIONNUMBER = "
    strSQL = strSQL & objSQLAssist.FormatString(strApplicationNumber) & " AND"
    strSQL = strSQL & " PAYMENTPROTECTIONSUBQUOTE.APPLICATIONFACTFINDNUMBER = "
    strSQL = strSQL & strApplicationFFNumber & " AND"
    strSQL = strSQL & " PAYMENTPROTECTIONSUBQUOTE.PPSUBQUOTENUMBER = "
    strSQL = strSQL & strPPSubQuoteNumber
        
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
    
    conn.Open objADOAssist.GetConnStr
    cmd.ActiveConnection = conn
            
    Set rstThisRecSet.Source = cmd
    
    rstThisRecSet.Open
    
    Set rstThisRecSet.ActiveConnection = Nothing
    

    'End - IVW
                
    If rstThisRecSet.EOF Then
        ' raise application error to be interpreted by calling object
        objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound
    End If
        
    Dim strPPSubQuoteClassDef As String
    Dim strPPDetailsClassDef As String
    
    'FIX ME use dom doc
    'When implemented as Dom Doc change ClassDef to IClassDef
    'Dim objIClassDef As IClassDef
    Dim objIomCMClassDef As omCMClassDef
    
    Set objIomCMClassDef = New omCMClassDef
    
    strPPSubQuoteClassDef = objIomCMClassDef.LoadPPSubQuoteData()

    strPPDetailsClassDef = objIomCMClassDef.LoadPPDetailsData()
            
    ' AD 05/10/01 Removed the MoveFirst as it's a ForwardOnly cursor. SYS2718.
    
    ' loop through the record set
    While Not rstThisRecSet.EOF
        strXML = objDOAssist.GetXMLFromRecordSet(rstThisRecSet, strPPSubQuoteClassDef)
        Set xmlDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
        Set xmlSubQuoteElem = xmlOut.appendChild(xmlDoc.documentElement)
        strXML = objDOAssist.GetXMLFromRecordSet(rstThisRecSet, strPPDetailsClassDef)
        Set xmlDoc = objXmlAssist.load(strXML, TypeName(Me), strFunctionName)
        Set xmlNode = xmlSubQuoteElem.appendChild(xmlDoc.documentElement)
        rstThisRecSet.MoveNext
    Wend
        
    Set cmd = Nothing
        rstThisRecSet.Close
        Set rstThisRecSet = Nothing
        conn.Close
    Set conn = Nothing
    
    
    GetSubQuoteDetails = xmlOut.xml
    
    Set objIomCMClassDef = Nothing
    Set objErrAssist = Nothing
    Set xmlDoc = Nothing
    Set xmlOut = Nothing
    Set objADOAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlIn = Nothing
    Set xmlNode = Nothing
    Set xmlSubQuoteElem = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
            
    Exit Function

GetSubQuoteDetailsVbErr:
        
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
        
    Set xmlDoc = Nothing
    Set xmlOut = Nothing
    Set objADOAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlIn = Nothing
    Set xmlNode = Nothing
    Set xmlSubQuoteElem = Nothing
    
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

    Const strFunctionName As String = "FindList"
        
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim objSQLAssist As New SQLAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    
    Dim strSQL As String
    Dim strXML As String
    
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlDoc As New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlSubQuoteElem As IXMLDOMElement
    Dim xmlListElem As IXMLDOMElement
    Dim xmlSubQuoteDefDoc As FreeThreadedDOMDocument40
    Dim xmlDetailsDefDoc As FreeThreadedDOMDocument40
    
    Dim objIADOAssist As IADOAssist
    Dim objIDOAssist As IDOAssist
    
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    Dim strApplicationNumber As String
    Dim strApplicationFFNumber As String
    
    strApplicationNumber = _
        objXmlAssist.GetTagValue(xmlIn.documentElement, "APPLICATIONNUMBER")
    strApplicationFFNumber = _
        objXmlAssist.GetTagValue(xmlIn.documentElement, "APPLICATIONFACTFINDNUMBER")
    
    If Len(strApplicationNumber) = 0 Or Len(strApplicationFFNumber) = 0 Then
        objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeMissingPrimaryTag, _
                                "APPLICATIONNUMBER = " & strApplicationNumber & _
                                "; APPLICATIONFACTFINDNUMBER = " & strApplicationFFNumber
    End If
    
    If objContext Is Nothing Then
        Set objIADOAssist = New ADOAssist
        Set objIDOAssist = New DOAssist
    Else
        Set objIADOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
        Set objIDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    Set xmlListElem = xmlOut.createElement(cstrSUBQUOTE_NAME & "LIST")
    xmlOut.appendChild xmlListElem
    
    strSQL = "SELECT * FROM PAYMENTPROTECTIONSUBQUOTE, PAYMENTPROTECTIONDETAILS WHERE"
    strSQL = strSQL & " PAYMENTPROTECTIONSUBQUOTE.APPLICATIONNUMBER = "
    strSQL = strSQL & " PAYMENTPROTECTIONDETAILS.APPLICATIONNUMBER AND"
    strSQL = strSQL & " PAYMENTPROTECTIONSUBQUOTE.APPLICATIONFACTFINDNUMBER = "
    strSQL = strSQL & " PAYMENTPROTECTIONDETAILS.APPLICATIONFACTFINDNUMBER AND"
    strSQL = strSQL & " PAYMENTPROTECTIONSUBQUOTE.PPSUBQUOTENUMBER = "
    strSQL = strSQL & " PAYMENTPROTECTIONDETAILS.PPSUBQUOTENUMBER AND"
    strSQL = strSQL & " PAYMENTPROTECTIONSUBQUOTE.APPLICATIONNUMBER = "
    strSQL = strSQL & objSQLAssist.FormatString(strApplicationNumber) & " AND"
    strSQL = strSQL & " PAYMENTPROTECTIONSUBQUOTE.APPLICATIONFACTFINDNUMBER = "
    strSQL = strSQL & strApplicationFFNumber
        
   'IVW -Disconnect record sets more efficiently
    
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
    

    'IVW - End
    
    If rstThisRecSet.EOF Then
        ' raise application error to be interpreted by calling object
        objErrAssist.ThrowError TypeName(Me), strFunctionName, oeRecordNotFound
    End If
            
    Dim objIomCMClassDef As IomCMClassDef
    
    Set objIomCMClassDef = New omCMClassDef
    
    Set xmlSubQuoteDefDoc = objIomCMClassDef.LoadPPSubQuoteData()

    Set xmlDetailsDefDoc = objIomCMClassDef.LoadPPDetailsData()
            
    ' AD 05/10/01 Removed MoveFirst as it's a ForwardOnly cursor. SYS2718
    
    ' loop through the record set
    While Not rstThisRecSet.EOF
        Set xmlSubQuoteElem = objIDOAssist.GetXMLFromRecordSet(rstThisRecSet, _
                                                               xmlSubQuoteDefDoc, _
                                                               xmlListElem)
        
        objIDOAssist.GetXMLFromRecordSet rstThisRecSet, _
                                         xmlDetailsDefDoc, _
                                         xmlSubQuoteElem
        rstThisRecSet.MoveNext
    Wend
        
    
    FindList = xmlOut.xml
    
    Set cmd = Nothing
        rstThisRecSet.Close
        Set rstThisRecSet = Nothing
        conn.Close
    Set conn = Nothing
    
    
    Set objIomCMClassDef = Nothing
    Set objErrAssist = Nothing
    Set xmlOut = Nothing
    Set objIADOAssist = Nothing
    Set objIDOAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlIn = Nothing
    Set xmlSubQuoteElem = Nothing
    Set xmlListElem = Nothing
    Set xmlSubQuoteDefDoc = Nothing
    Set xmlDetailsDefDoc = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
            
    Exit Function

FindListVbErr:
        
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
        
    Set xmlOut = Nothing
    Set objIADOAssist = Nothing
    Set objIDOAssist = Nothing
    Set objXmlAssist = Nothing
    Set xmlIn = Nothing
    Set xmlSubQuoteElem = Nothing
    Set xmlListElem = Nothing
    Set xmlSubQuoteDefDoc = Nothing
    Set xmlDetailsDefDoc = Nothing
    
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

Public Sub UpdateSubQuoteDetails(ByVal vstrXMLRequest As String)
' header ----------------------------------------------------------------------------------
' description: Updates the Payment Protection Sub Quote Details consisting of the
'              Payment Protection Sub quote and the Payment Protection Details
' pass:
'   vstrXMLRequest  xml Request data stream containing data to be persisted
'------------------------------------------------------------------------------------------
On Error GoTo UpdateSubQuoteDetailsVbErr

    Const strFunctionName As String = "UpdateSubQuoteDetails"
   
    Dim objErrAssist As New ErrAssist
    Dim objXmlAssist As New XMLAssist
    Dim objDOAssist As DOAssist
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlElement As IXMLDOMElement
    Dim xmlSubQuoteElem As IXMLDOMElement
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
         
    Set xmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)

    Set xmlSubQuoteElem = xmlIn.getElementsByTagName(cstrSUBQUOTE_NAME).Item(0)
    
    If xmlSubQuoteElem Is Nothing Then
        objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeMissingPrimaryTag, _
                                cstrSUBQUOTE_NAME & " tag not found"
    End If
        
    If objContext Is Nothing Then
        Set objDOAssist = New DOAssist
    Else
        Set objDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    UpdatePaymentProtectionSubQuote xmlSubQuoteElem.xml
    objXmlAssist.ChangeNodeName xmlSubQuoteElem, cstrSUBQUOTE_NAME, "PAYMENTPROTECTIONDETAILS"
    UpdatePaymentProtectionDetails xmlSubQuoteElem.xml
    
    Set objErrAssist = Nothing
    Set objXmlAssist = Nothing
    Set objDOAssist = Nothing
    Set xmlIn = Nothing
    Set xmlElement = Nothing
    Set xmlSubQuoteElem = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Sub

UpdateSubQuoteDetailsVbErr:

    Set objXmlAssist = Nothing
    Set objDOAssist = Nothing
    Set xmlIn = Nothing
    Set xmlElement = Nothing
    Set xmlSubQuoteElem = Nothing
    
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

Private Sub CreatePaymentProtectionSubQuote(ByVal vstrXMLRequest As String)
' header ----------------------------------------------------------------------------------
' description:
'   create an instance of PaymentProtectionSubQuote
' pass:
'   vstrXMLRequest  xml Request data stream containing data to be persisted
'------------------------------------------------------------------------------------------
On Error GoTo CreatePaymentProtectionSubQuoteVbErr

    Const strFunctionName As String = "CreatePaymentProtectionSubQuote"
   
    Dim objErrAssist As New ErrAssist
    
    Dim objDOAssist As DOAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
                    
    If objContext Is Nothing Then
        Set objDOAssist = New DOAssist
    Else
        Set objDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    'FIX ME use dom doc
    'When implemented as Dom Doc change ClassDef to IClassDef
    'Dim objIClassDef As IClassDef
    Dim objIomCMClassDef As omCMClassDef
    
    Set objIomCMClassDef = New omCMClassDef

    objDOAssist.Create vstrXMLRequest, objIomCMClassDef.LoadPPSubQuoteData()
    
    Set objDOAssist = Nothing
    Set objErrAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Sub

CreatePaymentProtectionSubQuoteVbErr:

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

End Sub

Private Sub CreatePaymentProtectionDetails(ByVal vstrXMLRequest As String)
' header ----------------------------------------------------------------------------------
' description:
'   create an instance of CreatePaymentProtectionDetails
' pass:
'   vstrXMLRequest  xml Request data stream containing data to be persisted
'------------------------------------------------------------------------------------------
On Error GoTo CreatePaymentProtectionDetailsVbErr

    Const strFunctionName As String = "CreatePaymentProtectionDetails"
   
    Dim objErrAssist As New ErrAssist
    
    Dim objDOAssist As DOAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
                    
    If objContext Is Nothing Then
        Set objDOAssist = New DOAssist
    Else
        Set objDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    'FIX ME
    'When implemented as Dom Doc change ClassDef to IClassDef
    'Dim objIClassDef As IClassDef
    Dim objIomCMClassDef As omCMClassDef
    
    Set objIomCMClassDef = New omCMClassDef
    
    objDOAssist.Create vstrXMLRequest, objIomCMClassDef.LoadPPDetailsData()
    
    Set objDOAssist = Nothing
    Set objErrAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Sub

CreatePaymentProtectionDetailsVbErr:

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

End Sub

Private Function GetHighestSubQuoteNumber( _
                        ByVal strApplicationNumber As String, _
                        ByVal strApplicationFactFindNumber As String) As Long
' header ----------------------------------------------------------------------------------
' description:  get highest PP Sub Quote Number based on Application Number and
'               Application Fact Find Number
' pass: strApplicationNumber         Application number
'       strApplicationFactFindNumber Application fact find number
'BG SYS1945 13/02/01 - changed the way it sets activeconnection to nothing
'------------------------------------------------------------------------------------------
On Error GoTo GetHighestSubQuoteNumberVbErr

    Const strFunctionName As String = "GetHighestSubQuoteNumber"

    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
        
    Dim objErrAssist As New ErrAssist
    Dim objSQLAssist As New SQLAssist
    Dim objADOAssist As ADOAssist
    
    Dim strSQL As String
    
    ' Validate parameters
    If Len(strApplicationNumber) < 1 Then
        ' "Invalid parameter value: empty string received"
        objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeInvalidParameter, _
                                "APPLICATIONNUMBER = " & _
                                strApplicationNumber
    End If
    
    If Len(strApplicationFactFindNumber) < 1 Then
        ' "Invalid parameter value: empty string received"
        objErrAssist.ThrowError TypeName(Me), _
                                strFunctionName, _
                                oeInvalidParameter, _
                                "APPLICATIONFACTFINDNUMBER = " & _
                                strApplicationFactFindNumber
    End If
    
    If objContext Is Nothing Then
        objADOAssist = New ADOAssist
    Else
        Set objADOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".ADOAssist")
    End If

    strSQL = "SELECT MAX(PPSUBQUOTENUMBER) FROM"
    strSQL = strSQL & " PAYMENTPROTECTIONSUBQUOTE WHERE"
    strSQL = strSQL & " APPLICATIONNUMBER = "
    strSQL = strSQL & objSQLAssist.FormatString(strApplicationNumber)
    strSQL = strSQL & " AND APPLICATIONFACTFINDNUMBER = "
    strSQL = strSQL & strApplicationFactFindNumber
        
    'IVW -Disconnect record sets more efficiently
    
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
    
    conn.Open objADOAssist.GetConnStr
    cmd.ActiveConnection = conn
            
    Set rstThisRecSet.Source = cmd
    
    rstThisRecSet.Open
    
    Set rstThisRecSet.ActiveConnection = Nothing
    
 
    'IVW
  
    If IsNull(rstThisRecSet.Fields(0).Value) Then
        ' sequence number not yet set on this field
        GetHighestSubQuoteNumber = 0
    Else
        GetHighestSubQuoteNumber = rstThisRecSet.Fields(0).Value
    End If

    Set objErrAssist = Nothing
    Set objSQLAssist = Nothing
    Set objADOAssist = Nothing
    Set cmd = Nothing
        rstThisRecSet.Close
        Set rstThisRecSet = Nothing
    conn.Close
    Set conn = Nothing
        
    
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

GetHighestSubQuoteNumberVbErr:

    Set objSQLAssist = Nothing
    Set objADOAssist = Nothing
    
    ' Close the connection if still open
    If conn.State = adStateOpen Then
        conn.Close
    End If
    
    Set cmd = Nothing
    Set conn = Nothing
    Set rstThisRecSet = Nothing
    
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

Private Sub UpdatePaymentProtectionSubQuote(ByVal vstrXMLRequest As String)
' header ----------------------------------------------------------------------------------
' description:
'   update an instance of Payment Protection Sub Quote
' pass:
'   vstrXMLRequest  xml Request data stream containing data to be persisted
'------------------------------------------------------------------------------------------
On Error GoTo UpdatePaymentProtectionSubQuoteVbErr

    Const strFunctionName As String = "UpdatePaymentProtectionSubQuote"
   
    Dim objErrAssist As New ErrAssist
    
    Dim objDOAssist As DOAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
                    
    If objContext Is Nothing Then
        Set objDOAssist = New DOAssist
    Else
        Set objDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    'FIX ME use dom doc
    'When implemented as Dom Doc change ClassDef to IClassDef
    'Dim objIClassDef As IClassDef
    Dim objIomCMClassDef As omCMClassDef
    
    Set objIomCMClassDef = New omCMClassDef

    objDOAssist.Update vstrXMLRequest, objIomCMClassDef.LoadPPSubQuoteData()
    
    Set objDOAssist = Nothing
    Set objErrAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Sub

UpdatePaymentProtectionSubQuoteVbErr:

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

End Sub

Private Sub UpdatePaymentProtectionDetails(ByVal vstrXMLRequest As String)
' header ----------------------------------------------------------------------------------
' description:
'   update an instance of Payment Protection Details
' pass:
'   vstrXMLRequest  xml Request data stream containing data to be persisted
'------------------------------------------------------------------------------------------
On Error GoTo UpdatePaymentProtectionDetailsVbErr

    Const strFunctionName As String = "UpdatePaymentProtectionDetails"
   
    Dim objErrAssist As New ErrAssist
    
    Dim objDOAssist As DOAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
                    
    If objContext Is Nothing Then
        Set objDOAssist = New DOAssist
    Else
        Set objDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    'FIX ME
    'When implemented as Dom Doc change ClassDef to IClassDef
    'Dim objIClassDef As IClassDef
    Dim objIomCMClassDef As omCMClassDef
    
    Set objIomCMClassDef = New omCMClassDef

    objDOAssist.Update vstrXMLRequest, objIomCMClassDef.LoadPPDetailsData()
    
    Set objDOAssist = Nothing
    Set objErrAssist = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Sub

UpdatePaymentProtectionDetailsVbErr:

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

    Const strFunctionName As String = "GetData"
    
    Dim objErrAssist As New ErrAssist
    
                
    Dim objDOAssist As DOAssist
    
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
            
    If objContext Is Nothing Then
        Set objDOAssist = New DOAssist
    Else
        Set objDOAssist = objContext.CreateInstance(gstrBASE_COMPONENT & ".DOAssist")
    End If
    
    'FIX ME use dom doc
    'When implemented as Dom Doc change ClassDef to IClassDef
    'Dim objIClassDef As IClassDef
    Dim objIomCMClassDef As omCMClassDef
    
    Set objIomCMClassDef = New omCMClassDef

    GetData = objDOAssist.GetData(vstrXMLRequest, objIomCMClassDef.LoadPPSubQuoteData)

    ' TODO ************************************************************************
    '   Add Any derived data here
    ' *****************************************************************************
    
    Set objDOAssist = Nothing
    Set objErrAssist = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

GetDataVbErr:
        
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

