VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "PAFBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      PAFBO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling
'Description:   PAF Business Object
'Dependencies:  PAFDO
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'
' Prog  Date     Description
' PSC   06/08/99 Created.
' RF    16/02/00 AQR SYS0251: Moved to omBase component.
' DB    28/05/02 SYS4767 - MSMS to Core Integration
' INR   05/06/02 SYS4803 PAF search not working after Core Integration
'------------------------------------------------------------------------------------------
Option Explicit
Implements ObjectControl
Private m_objContext As ObjectContext
Private Const cstrTABLE_NAME = "REQUEST"
'------------------------------------------------------------------------------------------
'History:
'
' Prog  Date     Description
' INR   30/10/02 BMIDS00729/CC021 Changed PAF call from QAS to GB Address manager
' INR   15/11/02 BMIDS00960/CC021 Changed PAF call from GB Address to GB National Register
' DB    27/03/03 BM0488 When address not found stop PAF writing error to eventlog.
'------------------------------------------------------------------------------------------
'BBG Specific History:
'
' Prog  Date        Description
' MSla  22/10/2004  BBG1524 - Make error message more meaningful for invalid area code
'------------------------------------------------------------------------------------------
'Baseline Specific History:
'
' Prog  Date     Description
' MV    28/11/2003  CORE12 - Amended FindPAFAddress() to call either QASPro or GB Address Mangager
' AS    10/02/2006  CORE243 GlobalParameterDO running outside MTS.
'------------------------------------------------------------------------------------------



Public Function FindPAFAddress(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Finds the addresses that match the passed in criteria
' pass:         vstrXMLRequest  xml Request data stream containing data which identifies
'                               the address to search for
'
' return:       FindPAFAddress  xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS" and xml representation of data
'                               or: TYPE="SYSERR" and <ERROR> element
' Note!                         A search which finds nothing is still a SUCCESS, it is up to
'                               the client to look for any addresses.
' Raise Errors:
'------------------------------------------------------------------------------------------
On Error GoTo FindPAFAddressVbErr
    
    Dim strFunctionName As String
    strFunctionName = "FindPAFAddress"
    Dim objContext As ObjectContext
    Set objContext = GetObjectContext()
    Dim objXmlIn As New FreeThreadedDOMDocument40
    Dim objXmlOut As New FreeThreadedDOMDocument40
    Dim objXmlData As New FreeThreadedDOMDocument40
    Dim objXmlResponseElem As IXMLDOMElement
    Dim objXmlRequestElem As IXMLDOMElement
    Dim objXMLElement As IXMLDOMElement
    Dim objXmlNode As IXMLDOMNode
    Dim objXmlItemList As IXMLDOMNodeList
    Dim objXmlItemElem As IXMLDOMElement
    Dim objAddressElem As IXMLDOMElement
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    Dim objGlobalParameterDO  As GlobalParameterDO
    Dim objPAF As Object
    Dim intPAFEngine As Integer
    Dim strWrapperobject As String
    Dim strErrorNumber As String
    Dim strErrorText As String
    Dim strResponseXml As String
    Dim strXMLData As String
    Dim strUsingPickListKey As String
    Dim strQASKey As String
    Dim strPostCode As String
    Dim nItem As Integer
        
    Set objXmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    Set objXmlResponseElem = objXmlOut.createElement("RESPONSE")
    ' Perform validation on the parameters
    Validate (vstrXMLRequest)
    If Not m_objContext Is Nothing Then
        Set objGlobalParameterDO = m_objContext.CreateInstance(gstrBASE_COMPONENT & ".GlobalParameterDO")
    Else
        Set objGlobalParameterDO = New GlobalParameterDO
    End If
    intPAFEngine = objGlobalParameterDO.GetCurrentParameterByType("PAFEngine", "AMOUNT")
    If intPAFEngine = 1 Then   'QASPRO
        strWrapperobject = "omQASProWrapper.omQASPro"
    ElseIf intPAFEngine = 2 Then 'omGBMAIL
        strWrapperobject = "omGBNatRegWrapper.omGBNatReg"
    Else
         objErrAssist.RaiseError TypeName(Me), _
                                strFunctionName, _
                                omiga4InvalidParameter, _
                                " Global Parameter PAFEngine Value not found"
    End If
    Set objPAF = CreateObject(strWrapperobject)
    Set objXmlRequestElem = objXmlIn.getElementsByTagName("REQUEST").Item(0)
    'Call Object FindPAFAddress function
    If Not objXmlRequestElem Is Nothing Then
        strUsingPickListKey = objXmlRequestElem.firstChild.baseName
                
        'An initial search will have address info to search on. If we have
        'selected an item from a picklist, we will have an addresskey to search on.
        'Do we have an ADDRESSKEY, or just an address to search for?
        If (StrComp(strUsingPickListKey, "ADDRESSKEY", vbTextCompare) = 0) Then
            strXMLData = objPAF.FindAddressWithKey(objXmlRequestElem.xml)
        Else
            strXMLData = objPAF.FindAddress(objXmlRequestElem.xml)
        End If
        Set objXmlData = objXmlAssist.load(strXMLData, TypeName(Me), strFunctionName)
        ' if error then raise problem and return error.
        Set objXMLElement = objXmlData.getElementsByTagName("RESPONSE").Item(0)
        If (objXMLElement.getAttribute("TYPE") = "APPERR") Then
            objXmlResponseElem.setAttribute "TYPE", "APPERR"
            'MSla 22/10/2004 BBG1524
            strErrorText = objXmlData.getElementsByTagName("DESCRIPTION").Item(0).Text
            'MSla 22/10/2004 BBG1524 - End

            Set objXMLElement = objXmlData.getElementsByTagName("NUMBER").Item(0)
            strErrorNumber = objXMLElement.Text
            Err.Raise strErrorNumber, "FindAddress", strErrorText
        Else
            objXmlResponseElem.setAttribute "TYPE", "SUCCESS"
        End If
    Else
        objErrAssist.RaiseError TypeName(Me), _
                                strFunctionName, _
                                omiga4MissingPrimaryTag, _
                                cstrTABLE_NAME & " tag not found"
    End If
    'Get the Postcode from the QASPostcode key and set it as a new attribute to be
    'used in the picklist
    Set objAddressElem = objXmlData.selectSingleNode("//RESPONSE/ADDRESSDATA")
    Set objAddressElem = objAddressElem.cloneNode(False)
    Set objXmlItemList = objXmlData.selectNodes("//RESPONSE/ADDRESSDATA/ITEM")
    For nItem = 0 To (objXmlItemList.length - 1)
        Set objXmlItemElem = objXmlItemList.Item(nItem)
        strQASKey = objXmlItemElem.getAttribute("KEY")
        strPostCode = GetPostCodeFromQASKey(strQASKey)
        objXmlItemElem.setAttribute "POSTCODE", strPostCode
        objAddressElem.appendChild objXmlItemElem
    Next
    objXmlResponseElem.appendChild objAddressElem
    FindPAFAddress = objXmlResponseElem.xml
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
FindPAFAddressExit:
    
    Set objXmlIn = Nothing
    Set objXmlOut = Nothing
    Set objXmlData = Nothing
    Set objXmlResponseElem = Nothing
    Set objXmlRequestElem = Nothing
    Set objXMLElement = Nothing
    Set objXmlNode = Nothing
    Set objXmlItemList = Nothing
    Set objXmlItemElem = Nothing
    Set objAddressElem = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set objGlobalParameterDO = Nothing
    Set objPAF = Nothing
    Exit Function
FindPAFAddressVbErr:
    
    If objErrAssist.IsWarning = True Then
        '   add message element to response block
        Dim objXmlDoc As New FreeThreadedDOMDocument40
        objXmlDoc.setProperty "NewParser", True
        objXmlDoc.validateOnParse = False
        objXmlDoc.loadXML objErrAssist.FormatMessageNode
        Set objXmlNode = objXmlResponseElem.appendChild(objXmlDoc.documentElement)
        Set objXmlDoc = Nothing
        Resume Next
    End If
    FindPAFAddress = objErrAssist.CreateErrorResponse
        
    If Not objContext Is Nothing Then
        objContext.SetAbort
    End If
    '   go to clean-up section
    Resume FindPAFAddressExit
End Function
Private Sub Validate(ByVal vstrXMLRequest As String)
' header ----------------------------------------------------------------------------------
' description:  performs validation for this object
'               if hard-coded rules completed OK calls scripting interface to perform
'               user scripted rules as appropriate
' pass:         vstrXMLRequest  xml Request data stream containing details of action to be
'                               performed and data required
'                               This is the full request block as received from the client
'
' return:       n/a
' Raise Errors: all errors detected to be raised to calling function
'------------------------------------------------------------------------------------------
On Error GoTo ValidateVbErr
    Dim strFunctionName As String
    strFunctionName = "Validate"
    Dim objXmlIn As FreeThreadedDOMDocument40
    Dim objSearchElement As IXMLDOMElement
    Dim objParameterNode As IXMLDOMNode
    Dim objXmlAssist As New XMLAssist
    Dim objErrAssist As New ErrAssist
    'DB SYS4767 - MSMS Integration
    Dim objScriptInterface As ScriptInterface
    Set objScriptInterface = New ScriptInterface
    Dim blnParamsAreBlank As Boolean
    Dim nLenAttributes As Integer
    Dim nCounter As Integer
    nCounter = 0
    'DB End
    ' check at least one search criteria has been entered
    Set objXmlIn = objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    'DB SYS4767 - MSMS Integration
    Set objSearchElement = objXmlIn.getElementsByTagName("ADDRESS").Item(0)
    'DB End
    If objSearchElement Is Nothing Then
    'DB SYS4767 - MSMS Integration
    'Could be ADDRESS or ADDRESSKEY
        Set objSearchElement = objXmlIn.getElementsByTagName("ADDRESSKEY").Item(0)
        If objSearchElement Is Nothing Then
        objErrAssist.RaiseError TypeName(Me), _
                                strFunctionName, _
                                omiga4MissingPrimaryTag, _
                                    "ADDRESS or ADDRESSKEY tag not found"
        End If
    End If
    'DB End
    blnParamsAreBlank = True
    'DB SYS4767 MSMS Integration
    'Don't count the Picklist or searchkey attributes as valid search information
    'Check other attributes to see if we have entered search info
    nLenAttributes = objSearchElement.Attributes.length - 1
        
    Do While (nCounter < nLenAttributes) And blnParamsAreBlank = True
        If Len(objSearchElement.Attributes.Item(nCounter).xml) > 0 Then
            blnParamsAreBlank = False
        End If
        nCounter = nCounter + 1
    Loop
    'DB End
    If blnParamsAreBlank = True Then
        objErrAssist.RaiseError TypeName(Me), strFunctionName, omiga4err120
    End If
        
    objScriptInterface.RunScript App.Title, _
                                 TypeName(Me), _
                                 strFunctionName, _
                                 vstrXMLRequest
    '   clean-up code
    Set objScriptInterface = Nothing
    Set objXmlIn = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set objSearchElement = Nothing
    Set objParameterNode = Nothing
    Exit Sub
ValidateVbErr:
    '   clean-up code
    Set objScriptInterface = Nothing
    Set objXmlIn = Nothing
    Set objXmlAssist = Nothing
    Set objErrAssist = Nothing
    Set objSearchElement = Nothing
    Set objParameterNode = Nothing
    '   re-raise error to calling function
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
Function GetPostCodeFromQASKey(sPostCodeKeyIn) As String
' header ----------------------------------------------------------------------------------
' description:  Get the postcode from within the QAS Key, required only for display purposes
'
' pass:         sPostCodeKeyIn  QAS postcode key
'
' return:       sPostCode
' Raise Errors: all errors detected to be raised to calling function
'------------------------------------------------------------------------------------------
On Error GoTo GetPostCodeFromQASKeyVbErr
    Dim strFunctionName As String
    strFunctionName = "GetPostCodeFromQASKey"
    Dim objErrAssist As New ErrAssist
    Dim iPosFirst, iPosSecond As Integer
    Dim sPostCode As String
    iPosFirst = 1
    iPosSecond = 0
      
    iPosFirst = InStr(iPosFirst, sPostCodeKeyIn, ":", vbTextCompare) + 1
    If iPosFirst > 0 Then
        
        iPosSecond = InStr(iPosFirst, sPostCodeKeyIn, ":", vbTextCompare)
        If iPosSecond > 0 Then
            sPostCode = Mid(sPostCodeKeyIn, iPosFirst, iPosSecond - iPosFirst)
        Else
            objErrAssist.RaiseError TypeName(Me), strFunctionName, oeXMLInvalidAttributeValue
        End If
    Else
        objErrAssist.RaiseError TypeName(Me), strFunctionName, oeXMLInvalidAttributeValue
    End If
    GetPostCodeFromQASKey = sPostCode
    Set objErrAssist = Nothing
    Exit Function
GetPostCodeFromQASKeyVbErr:
    'Have had instances where QAS returns '/' within key instead of ':'
    'If errors just display the Postcodekey
    GetPostCodeFromQASKey = sPostCodeKeyIn
    Set objErrAssist = Nothing
End Function
Private Sub ObjectControl_Activate()
    Set m_objContext = GetObjectContext()
End Sub
Private Function ObjectControl_CanBePooled() As Boolean
ObjectControl_CanBePooled = False
End Function
Private Sub ObjectControl_Deactivate()
Set m_objContext = Nothing
End Sub
