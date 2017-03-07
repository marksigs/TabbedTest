VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "BankWizardBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' TODO ************************************************************************
'   replace all instances of 'Omiga4BO' with actual class name
' *****************************************************************************

'Workfile:      BankWizardBO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling

'Description:   Code template for omiga4 Business Object which 'supports transactions'
'               only
'               i.e. any code that does not require transaction support will reside in
'               this first level module. Any functionality that does require
'               transaction support will be delegated to an appropriate (second level)
'               module (Omiga4TxBO in this example)
'
'Dependencies:  List any other dependent components
'               e.g. Omiga4TxBO, Omiga4DO
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'IK     30/06/99    Created
'RF     30/09/99    Applied changes raised by code review of 30/09/99, including:
'                   removed AnonInterfaceFunction
'                   removed calls to Validate
'                   improved error handling
'RF     04/10/99    Added profiling
'PSC    21/10/99    Amend warning processing
'PSC    14/12/99    Amend to use multi[le interfaces. The default interface is used for
'                   scripting and cross component calls. The IBankWizardBO interface is for
'                   use internal to a component
'PSC    16/12/99    Amend interfaces
'PSC    21/12/99    Amend so that xmlOUT in New'd as a DOMDocument in FindList and
'                   GetData
'PSC    01/02/00    Make objContext, objErrAssist and objXmlAssist object level rather
'                   than procedure level
'                   Don't clone the response node
'MDC    23/05/00    Add validation to ensure Sortcode and Account Number exist, they are
'                   numeric and that the sort code does not contain hyphens.
'IVW    16/07/00    Removed the DO - The External VC Component is now the DO.
'IVE    16/07/00    Created GetAllBankDetails from original DO.
'DRC    3/10/01     SYS2745 Replaced .SetAbort with .SetComplete in Get, Find & Validate Methods
'TJ     13/01/02    SYS3745 Removed comments causing errors in OOSS
'SG     27/05/02    SYS4767 MSMS to Core integration
'------------------------------------------------------------------------------------------

Option Explicit

' TODO ************************************************************************
'   Add interface definition class and amend this to implement it
' *****************************************************************************
Implements IBankWizardBO
Implements ObjectControl

'=============================================
'Constant Declaration Section
'=============================================
' TODO ************************************************************************
'   replace NODENAME with actual node name e.g. QUOTATION
' *****************************************************************************
Private Const cstrROOT_NODE_NAME = "BANKDETAILS"

'=============================================
'Variable Declaration Section
'=============================================
Private m_objContext As ObjectContext
Private m_objErrAssist As ErrAssist
Private m_objXmlAssist As XMLAssist

Public Function GetBankDetails(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo GetBankDetailsVbErr

    Const strFunctionName As String = "GetBankDetails"
        
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                    
    ' Run validation checks
    Set xmlTempResponseNode = IBankWizardBO_ValidateBankDetails(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
                                    
    ' Delegate to DomDocument based method and attach returned data to our response
    'SG 27/05/02 SYS4767
    'Set xmlTempResponseNode = IBankWizardBO_GetBankDetails(xmlIn.documentElement)
    Set xmlTempResponseNode = IBankWizardBO_GetBankWizardDetails(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
        
    GetBankDetails = xmlResponseElem.xml
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

GetBankDetailsExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing

    Exit Function

GetBankDetailsVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    GetBankDetails = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetBankDetailsExit

End Function
   
Private Function IBankWizardBO_GetBankDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Get a single instance of the persistant data associated with this
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo GetBankDetailsVbErr
    
    Const strFunctionName As String = "IBankWizardBO_GetBankDetails"

    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
  '  Dim objIBankWizardDO As IBankWizardDO
               
    'Need Std Response if this is called directly - bypassed by calling BO points to reqd node
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
      
'    If m_objContext Is Nothing Then
 '       Set objIBankWizardDO = New BankWizardDO
 '   Else
 '       Set objIBankWizardDO = m_objContext.CreateInstance(App.Title & ".BankWizardDO")
 '   End If
                
    If vxmlRequest.nodeName = cstrROOT_NODE_NAME Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName(cstrROOT_NODE_NAME).Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrROOT_NODE_NAME & " tag not found"
    End If
                
    ' Set xmlDataNode = objIBankWizardDO.GetBankDetails(xmlRequestNode)
    Set xmlDataNode = GetAllBankDetails(xmlRequestNode)
    ' BG SYS1277 25/08/00
    If Not xmlDataNode Is Nothing Then
    xmlResponseElem.appendChild xmlDataNode
    End If
    
    Set IBankWizardBO_GetBankDetails = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
GetBankDetailsExit:
    
    Set xmlRequestNode = Nothing
    'Set objIBankWizardDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Exit Function

GetBankDetailsVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IBankWizardBO_GetBankDetails = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume GetBankDetailsExit

End Function


Private Function IBankWizardBO_ValidateBankDetails(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  performs hard coded validation for this object
' pass:         vxmlRequest  xml Request data stream containing details of action to be
'                            performed and data required
'                            This is the full request block as received from the client
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo IBankWizardBO_ValidateBankDetailsVbErr
    
    Const strFunctionName As String = "IBankWizardBO_ValidateBankDetails"

    Dim xmlRequestTypeNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim objXmlSortCode As IXMLDOMNode
    Dim objNode As IXMLDOMNode
      
    Dim strSortCode As String
    Dim strAccountNumber As String
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    'Ensure that SORTCODE and ACCOUNTNUMBER tags are present
    
    'SG 27/05/02 SYS4767 START
    'strSortCode = m_objXmlAssist.GetTagValue(vxmlRequest, "SORTCODE")
    'strAccountNumber = m_objXmlAssist.GetTagValue(vxmlRequest, "ACCOUNTNUMBER")
    Dim objAddressElem As IXMLDOMElement
    
    Set objAddressElem = vxmlRequest.selectSingleNode("//REQUEST/BANKDETAILS")

    strSortCode = objAddressElem.getAttribute("SORTCODE")
    strAccountNumber = objAddressElem.getAttribute("ACCOUNTNUMBER")
    'SG 27/05/02 SYS4767 END
    
    If strSortCode = "" Or strAccountNumber = "" Then
        'Raise error
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 544, , "Bank Sort Code and Account Number must be specified"
    End If
    
    If InStr(strSortCode, "-") Then
        'Bank Wizard RSO requires sort code in the form 999999 i.e. no hyphens
        strSortCode = RemoveHyphens(strSortCode)
        
        'SG 27/05/02 SYS4767 START
        'Set objNode = vxmlRequest.ownerDocument.selectSingleNode("REQUEST/BANKDETAILS")
        'Set objXmlSortCode = m_objXmlAssist.GetMandatoryNode(objNode, "SORTCODE")
        'objXmlSortCode.Text = strSortCode
        objAddressElem.setAttribute "SORTCODE", strSortCode
        'SG 27/05/02 SYS4767 END
    End If
    
    'Validate that SORTCODE is numeric
    If Not IsNumeric(strSortCode) Then
        'Raise error
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 544, , "Bank Sort Code is not numeric"
    End If
    
    'Validate that ACCOUNTNUMBER is numeric
    If Not IsNumeric(strAccountNumber) Then
        'Raise error
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, 544, , "Bank Account Number is not numeric"
    End If
    
    
    '-----------------------------------------------------------------------------
    ' hard-coded validation OK
    ' return the response
    '-----------------------------------------------------------------------------
    
    Set IBankWizardBO_ValidateBankDetails = xmlResponseElem.cloneNode(True)
    
IBankWizardBO_ValidateBankDetailsExit:

    Set xmlRequestTypeNode = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set objXmlSortCode = Nothing
    Set objNode = Nothing
    Exit Function

IBankWizardBO_ValidateBankDetailsVbErr:
    
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IBankWizardBO_ValidateBankDetails = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    ' go to clean-up section
    Resume IBankWizardBO_ValidateBankDetailsExit

End Function
'SG 2705/02 SYS4647 Added Function
'------------------------------------------------------------------------
' Procedure GetBankWizardDetails
' Purpose:
'                   Takes the sort code provided and returns
'                   the associated branch address details from
'                   Bank Wizard. If an account number is
'                   specified, this is also validated.
' Input parameters: vxmlRequestNode holding XML request
'
' Output parameters: xmlResponseElem which will hold the XML response
'------------------------------------------------------------------------
Private Function IBankWizardBO_GetBankWizardDetails(ByVal vxmlRequestNode As IXMLDOMElement) As IXMLDOMNode

    '<!ELEMENT REQUEST (BANKDETAILS)>
    '  <!ATTLIST REQUEST USERID CDATA #REQUIRED>
    '  <!ATTLIST REQUEST UNITID CDATA #REQUIRED>
    '  <!ATTLIST REQUEST MACHINENAME CDATA #REQUIRED>
    '  <!ATTLIST REQUEST CHANNELID CDATA #REQUIRED>
    '  <!ATTLIST REQUEST USERAUTHORITYLEVEL CDATA #REQUIRED>
    '  <!ATTLIST REQUEST OPERATION CDATA #FIXED "ValidateBankDetails">
    '  <!ELEMENT BANKDETAILS EMPTY>
    '    <!ATTLIST BANKDETAILS SORTCODE CDATA #REQUIRED>
    '    <!ATTLIST BANKDETAILS ACCOUNTNUMBER CDATA #IMPLIED>

Dim xdcBankWizardDoc As FreeThreadedDOMDocument40

Dim objBankWizard As omBankWizardWrapper.omBankWizard
Dim xelBankDetails As IXMLDOMElement
Dim xelNewResponse As IXMLDOMElement
Dim xmlResponseElem As IXMLDOMElement
Dim xmlOut As New FreeThreadedDOMDocument40
xmlOut.validateOnParse = False
xmlOut.setProperty "NewParser", True
Dim xelNewChild As IXMLDOMElement
Const strFunctionName As String = "IBankWizardBO_GetBankWizardDetails"
Dim strTemp As String
Dim astrAddLines() As String

On Error GoTo GetBankWizardDetailsExit

    ' Set default return state (if in MTS)
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    'Call the BankWizard API to retrieve the bank details
    Set xdcBankWizardDoc = New FreeThreadedDOMDocument40
    xdcBankWizardDoc.validateOnParse = False
    xdcBankWizardDoc.setProperty "NewParser", True
    
    ' Add validation to check if Bank Wizard itself is available
On Error Resume Next
    Set objBankWizard = CreateObject("omBankWizardWrapper.omBankWizard")
    If objBankWizard Is Nothing Then
        errThrowError strFunctionName, oeObjectNotCreatable, "Bank Wizard Not Available"
    End If
On Error GoTo GetBankWizardDetailsExit

    Set xdcBankWizardDoc = xmlLoad(objBankWizard.CheckDetails(vxmlRequestNode.xml), strFunctionName)
    
    
    If xdcBankWizardDoc.selectSingleNode("RESPONSE/@TYPE").Text = "BANKWIZARDERROR" Then
        Select Case xdcBankWizardDoc.selectSingleNode("RESPONSE/ERROR/NUMBER").Text
        Case 2
            errThrowError strFunctionName, oeRecordNotFound, "Sort-code not allocated to any bank branch"
        Case 4
            errThrowError strFunctionName, oeRecordNotFound, "Unknown Account Number Specified"
        Case 32
            errThrowError strFunctionName, oeRecordNotFound, "Account number format incorrect"
        Case 64
            errThrowError strFunctionName, oeRecordNotFound, "Sort-code format incorrect"
        Case Else
            errThrowError strFunctionName, oeRecordNotFound, "Unidentified error from Bank Wizard"
        End Select
        
    ElseIf xdcBankWizardDoc.selectSingleNode("RESPONSE[@TYPE='BANKWIZARDWARNING']") Is Nothing Then
        errCheckXMLResponse xdcBankWizardDoc.xml, True
    End If
    
    'Get BANK details from the new response
    Set xelBankDetails = xmlGetMandatoryNode(xdcBankWizardDoc, "RESPONSE/BANKDATA")

    ' add PROVIDER node to response
    Set xelNewResponse = xmlResponseElem.ownerDocument.createElement("BANKDATA")
    xmlResponseElem.appendChild xelNewResponse

    'Set the relevant attributes from BANKDATA on the BANK element
    xelNewResponse.setAttribute "ACTIVETO", xelBankDetails.getAttribute("DELETEDDATE")
    xelNewResponse.setAttribute "BRANCHNAME", xelBankDetails.getAttribute("BRANCHTITLE")
    xelNewResponse.setAttribute "BANKSORTCODE", xelBankDetails.getAttribute("SORTCODE")
    xelNewResponse.setAttribute "VALIDATEDINBANKWIZARD", "1"

    'Create a name node under the bank node and set the attributes here
    Set xelNewChild = xmlResponseElem.ownerDocument.createElement("ORGANISATIONNAME")
    xelNewResponse.appendChild xelNewChild
    xelNewChild.setAttribute "FULLNAME", xelBankDetails.getAttribute("SHORTNAME")

    'Create an address node under the bank node and set the attributes here
    Set xelNewChild = xmlResponseElem.ownerDocument.createElement("ADDRESS")
    xelNewResponse.appendChild xelNewChild

    'Copy the values from the bank wizard response
    'INR Have to compromise a bit here as the number is included in the first addressline.
    xelNewChild.setAttribute "ADDRESSLINE1", xelBankDetails.getAttribute("ADDRESS1")
    xelNewChild.setAttribute "POSTCODE", xelBankDetails.getAttribute("POSTCODE")

    'Assign these values to the response
    xelNewChild.setAttribute "ADDRESSLINE2", xelBankDetails.getAttribute("ADDRESS2")
    xelNewChild.setAttribute "ADDRESSLINE3", xelBankDetails.getAttribute("TOWN")
    xelNewChild.setAttribute "ADDRESSLINE4", xelBankDetails.getAttribute("COUNTY")
    
    Set IBankWizardBO_GetBankWizardDetails = xmlResponseElem

GetBankWizardDetailsExit:
    'Free any created objects
    Set objBankWizard = Nothing
    Set xelBankDetails = Nothing
    Set xdcBankWizardDoc = Nothing

    errCheckError strFunctionName, TypeName(Me)
End Function

Private Sub ObjectControl_Activate()

    Set m_objContext = GetObjectContext()
    Set m_objErrAssist = New ErrAssist
    Set m_objXmlAssist = New XMLAssist

End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    
    ObjectControl_CanBePooled = False
    
End Function

Private Sub ObjectControl_Deactivate()

    Set m_objContext = Nothing
    Set m_objErrAssist = Nothing
    Set m_objXmlAssist = Nothing

End Sub

Private Function RemoveHyphens(ByVal strTextIn As String) As String
'Removes hyphens from a text string

On Error GoTo RemoveHyphensVbErr

Dim nPos As Integer
Dim nLen As Integer
Dim strTextOut As String
    
    Const strFunctionName As String = "RemoveHyphens"

    If Len(Trim$(strTextIn)) = 0 Then
        RemoveHyphens = strTextIn
        Exit Function
    End If
    
    strTextOut = strTextIn
    Do
        nLen = Len(strTextOut)
        nPos = InStr(strTextOut, "-")
        If nPos Then
            If nPos = 1 Then
                strTextOut = Right$(strTextOut, nLen - 1)
            ElseIf nPos = nLen Then
                strTextOut = Left$(strTextOut, nLen - 1)
            Else
                strTextOut = Left$(strTextOut, nPos - 1) & Right$(strTextOut, nLen - nPos)
            End If
        End If
    Loop Until nPos = 0

    RemoveHyphens = strTextOut
    Exit Function
    
RemoveHyphensVbErr:
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    'Raise the error for the calling method to deal with appropriately
    Err.Raise Err.Number, Err.Source, Err.Description
    Exit Function
    
End Function

Private Function GetAllBankDetails(ByVal vxmlTableElement As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Get the data for a single instance of the persistant data associated with
'   this data object
' pass:
'   vxmlTableElement  xml element containing the request
' return:                xml node containing retrieved data
'------------------------------------------------------------------------------------------
On Error GoTo GetAllBankDetailsVbErr

    Const strFunctionName As String = "GetAllBankDetails"
      
    Dim strValidationGenOut As String
    Dim strElementText As String
    Dim xmlValidSORTOut As FreeThreadedDOMDocument40
    Dim xmlValidACCOUNTOut As FreeThreadedDOMDocument40
    Dim xmlReturnBlock As New FreeThreadedDOMDocument40
    xmlReturnBlock.validateOnParse = False
    xmlReturnBlock.setProperty "NewParser", True
    Dim xmlTableClone As IXMLDOMElement
    Dim xmlSORTErrNoElement As IXMLDOMElement
    Dim xmlACCOUNTErrNoElement As IXMLDOMElement
    Dim intNodesLength As Integer
    Dim strTempElem As String
    Dim xmlRoot As IXMLDOMElement
    Dim xmlReturnNode As IXMLDOMNode
    Dim xmlSortCodeNode As IXMLDOMNode
    Dim xmlAccountNumberNode As IXMLDOMNode
    Dim xmlChildElement As IXMLDOMElement
    Dim strErrorNumber As String
    Dim strSortNoErrorText As String
    Dim strAccountNoErrorText As String
                
    Set xmlChildElement = xmlReturnBlock.createElement("BANKWIZARD")
    Set xmlReturnNode = xmlReturnBlock.appendChild(xmlChildElement)
    
    Set xmlChildElement = xmlReturnBlock.createElement("SORTCODECHECK")
    Set xmlSortCodeNode = xmlReturnNode.appendChild(xmlChildElement)
    
    Set xmlChildElement = xmlReturnBlock.createElement("ACCOUNTNUMBERCHECK")
    Set xmlAccountNumberNode = xmlReturnNode.appendChild(xmlChildElement)
        
    ' SYS3153  - TJ 04/12/01    Changed to late binding
    Dim objBankDetailsInterface As Object
    
    Set objBankDetailsInterface = CreateObject("BankDetailsInterface.BankDetailsInterface")
    ' End change
    
    Set xmlTableClone = vxmlTableElement.cloneNode(True)
    
    m_objXmlAssist.ChangeNodeName xmlTableClone, "BANKDETAILS", "BANKDETAILSINPUT"
    
    '1st VALIDATION CALL Checks SORT CODE's 6 digits(NO hyphens) authenticity
    
    strValidationGenOut = objBankDetailsInterface.ValidateSortCode(xmlTableClone.xml)
    Set xmlValidSORTOut = m_objXmlAssist.load(strValidationGenOut, TypeName(Me), strFunctionName)
    Set xmlSORTErrNoElement = xmlValidSORTOut.getElementsByTagName("STATUS").Item(0)
    
    ' If an error has occurred retrieve error string from the bank wizard component XML error block.
    ' This error must be serious because if there is no Status Tag the BankWizard returns a
    ' dedicated error XML block.
          
    If xmlSORTErrNoElement Is Nothing Then ' no status block - serious error
        
        Set xmlSORTErrNoElement = xmlValidSORTOut.getElementsByTagName("BANKDETAILSERRORTEXT").Item(0)
        
        strElementText = xmlSORTErrNoElement.Text
        
        If xmlSORTErrNoElement Is Nothing Then
                m_objErrAssist.ThrowError TypeName(Me), _
                                  strFunctionName, _
                                  544, _
                                  , "Sort Code: Missing STATUS and ERROR DETAILS elements"
        Else
                m_objErrAssist.ThrowError TypeName(Me), _
                                  strFunctionName, _
                                  544, _
                                  , strElementText
        End If
      
    Else
        ' check the status field return to ensure no error
        
        strElementText = xmlSORTErrNoElement.Text
        
        strErrorNumber = "0"
            
        Select Case Trim(strElementText)
        
            Case "2" ' Error in the sortcode block
                strErrorNumber = "544"
            Case "1" ' Warning in the sortcode block
                strErrorNumber = "545"
                    
        End Select
        
        If strErrorNumber <> "0" Then
            ' get the error text from the return block
            strSortNoErrorText = xmlValidSORTOut.getElementsByTagName("MESSAGETEXT").Item(0).Text
            m_objErrAssist.ThrowError TypeName(Me), strFunctionName, CLng(strErrorNumber), , strSortNoErrorText
        End If
            
    End If
    
    ' Remove trailing spaces from the returned strings.
    Set xmlRoot = xmlValidSORTOut.documentElement
        
    intNodesLength = xmlRoot.childNodes.length
    
    For intNodesLength = 0 To intNodesLength - 1
        strTempElem = Trim(xmlRoot.childNodes.Item(intNodesLength).Text)
        xmlRoot.childNodes.Item(intNodesLength).Text = strTempElem
    Next
    
    xmlSortCodeNode.appendChild xmlValidSORTOut.documentElement
    
    ' If the sort code is okay then we can check the account number.
    ' The MESSAGETEXT tag will contain either SUCCESS if everything was okay, or if there
    ' are problems it will contain an error string.
    ' The STATUS tag will contain either 0 - Success, 1 - Warning, 2 - Error
    
    If xmlSORTErrNoElement.Text = "0" Then
        ' Only if the sort code validation was successful can we do an account number validation
        strValidationGenOut = objBankDetailsInterface.ValidateAccountDetails(xmlTableClone.xml)
        Set xmlValidACCOUNTOut = m_objXmlAssist.load(strValidationGenOut, TypeName(Me), strFunctionName)
        Set xmlACCOUNTErrNoElement = xmlValidACCOUNTOut.getElementsByTagName("STATUS").Item(0)
        
        If xmlACCOUNTErrNoElement Is Nothing Then
            Set xmlACCOUNTErrNoElement = xmlValidACCOUNTOut.getElementsByTagName("BANKDETAILSERRORTEXT").Item(0)
            strElementText = xmlACCOUNTErrNoElement.Text
        
            If xmlACCOUNTErrNoElement Is Nothing Then
                m_objErrAssist.ThrowError TypeName(Me), _
                                  strFunctionName, _
                                  544, _
                                  , "Account Code: Missing STATUS and ERROR DETAILS elements"
            Else
                m_objErrAssist.ThrowError TypeName(Me), _
                                  strFunctionName, _
                                  544, _
                                  , strElementText
            End If
      
        Else
            ' check the status field return to ensure no error
        
            strElementText = xmlACCOUNTErrNoElement.Text
        
            strErrorNumber = "0"
            
            Select Case Trim(strElementText)
       
                Case "2" ' Error in the account number block
                    strErrorNumber = "544"
                Case "1" ' Warning in the account number block
                    strErrorNumber = "545"
                    
            End Select
        
            If strErrorNumber <> "0" Then
                ' get the error text from the return block
                strAccountNoErrorText = xmlValidACCOUNTOut.getElementsByTagName("MESSAGETEXT").Item(0).Text
                m_objErrAssist.ThrowError TypeName(Me), strFunctionName, CLng(strErrorNumber), , strAccountNoErrorText
            End If
        End If
    End If
        
    xmlAccountNumberNode.appendChild xmlValidACCOUNTOut.documentElement
        
    Set GetAllBankDetails = xmlReturnBlock.documentElement
       
    Set xmlValidSORTOut = Nothing
    Set xmlValidACCOUNTOut = Nothing
    Set xmlReturnBlock = Nothing
    Set xmlTableClone = Nothing
    Set xmlSORTErrNoElement = Nothing
    Set xmlACCOUNTErrNoElement = Nothing
    Set xmlRoot = Nothing
    Set xmlReturnNode = Nothing
    Set xmlSortCodeNode = Nothing
    Set xmlAccountNumberNode = Nothing
    Set xmlChildElement = Nothing
    
    Set objBankDetailsInterface = Nothing
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

    Exit Function

GetAllBankDetailsVbErr:
    
    Set xmlValidSORTOut = Nothing
    Set xmlValidACCOUNTOut = Nothing
    Set xmlReturnBlock = Nothing
    Set xmlTableClone = Nothing
    Set xmlSORTErrNoElement = Nothing
    Set xmlACCOUNTErrNoElement = Nothing
    Set xmlRoot = Nothing
    Set xmlReturnNode = Nothing
    Set xmlSortCodeNode = Nothing
    Set xmlAccountNumberNode = Nothing
    Set xmlChildElement = Nothing
    Set objBankDetailsInterface = Nothing
            
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



