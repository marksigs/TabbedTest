VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "AdminInterfaceBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      AdminInterfaceBO.cls
'Copyright:     Copyright © 2001 Marlborough Stirling

'Description:   Administration System Interface Object.
'               Handles requests to the administration system

'Dependencies:  Add any other dependent components
'
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'PSC    05/01/01    Created
'PSC    27/02/01    SYS1751 Amend customer numbers and roles coming back from GetAccountDetails
'PSC    27/02/01    SYS1751 Amend GetCustomerDetails to return different details based on
'                   Customer Number passed in
'PSC    27/02/01    SYS1751 Amend to return channel id based on the channel id in the request
'PSC    28/02/01    SYS1751 Amend GetCustomerDetails to return the customer number based on the
'                   Customer Number passed in
'SR     05/06/01    SYS2325 New method 'GetNewNumbers'
'SR     13/06/01    SYS2362 Modified methods 'FindCustomer' and 'FindBusinessForCustomer'
'RF     16/11/01    SYS2927 (Omiplus CC34) Added ValidateUserLogon and LogOffUser.
'RF     13/12/01    SYS3489 (Omiplus CC32) ProcessCompletionsInterface now calls omToCBA (via OmiQueue)
'DS     18/12/01    SYS2325 (Omiplus CC02) Added call into ODI for getting customer//account numbers.
'MDC    07/02/2002  SYS4020 - Make AdminRules late bound
' DM  18/02/2002    SYS4105 changed XML to match specification and the attribute within GetNewNumbers.
'DS     20/02/2002  SYS2721 Add function GetNewNumberWithAutoLogon
'DM 20/02/2002  SYS2712 change call from getnewnumbers to getnewnumberswithautologon.
'DS     22/03/2002  SYS4304 Fixed a problem with findCustomer
'DS     26/03/02    SYS4323 Improved error handling
'DM     02/04/02    SYS4350 Implement SaveThirdpartyDetails
'JLD    22/04/02    SYS3552 send all payment types to be interfaces except ROF and Cancel Balance
'DS     30/04/02    Use FreeThreadedDOMDocument40.
'STB    20/05/02    SYS4609 Implemented Get/SaveUserMortgageAdminDetails and SaveMortgageProductDetails.
'------------------------------------------------------------------------------------------
Option Explicit

Implements ObjectControl

Private m_objContext As ObjectContext

'------------------------------------------------------------------------------------------
'BMIDS Specific History:
'
'Prog   Date        Description
'GHun   23/05/02    BMIDS00005 Added functions GetCRSContactData and UpdateCRSContactLog
'MDC    06/06/2002  BMIDS00025 Added function GetIntroducerData
'MV     21/06/2002  BMIDS00077 Core Upgrade - 7.0.2- Amended DORequest()
'GHun   10/06/2002  BMIDS00026 IWP2 BM060 Added function UpdateCRSCustomer
'GHun   13/06/2002  BMIDS00025, BMIDS00026, BMID00056 Added function CallBMidsAdminSystem &
'                   changed all other functions that need to access the BMids admin system
'                   to use the new function instead.
'GHun   26/06/2002  BMIDS00092 Save BMids admin system XML input and output for debugging
'MDC    16/07/2002  BMIDS00006 CAWP1 BM054 - Add call through for GetAccountCustomers
'GHun   06/08/2002  BMIDS00006 CAWP1 Changed GetAccountDetails to call BMIDS admin system
'MDC    20/08/2002  BMIDS00336 CCWP1 Added ExperianBureau & ExperianCreditCheck methods
'                   to call through to BMIDS
'PSC    20/09/2002  BMIDS00480 CIWP1 BM096 - Change ProcessCompletionsInterface to get the
'                   data using a schema name from GlobalParameters
'MO     21/10/2002  BMIDS00663, Made changes to delete the component omInt and place the
'                   fucntionality into the omAdmin Component
'PSC    22/10/2002  BMIDS00676  Add RunGenericInterface
'RF     10/06/2003  00008119    Add extra logging to CallBMidsAdminSystem
'RF     17/07/2003  00008267    Extend CallBMidsAdminSystem logging
'SR     17/08/2004  BMIDS815 - New method getAccountRefresh
'------------------------------------------------------------------------------------------

'------------------------------------------------------------------------------------------
'MARS Specific History:
'
'Prog   Date        Description
'SD     16/08/2005  Mar47 Added globalparameters and Combo
'PSC    19/10/2005  MAR265 Amend CallBMidsAdminSystem to CallExternalSystem
'PSC    31/10/2005  MAR352 Amend to call new admin interface component
'PSC    16/11/2005  MAR601 Correct attribute name for debugging of CallExternalSystem

'------------------------------------------------------------------------------------------
'EPSOM Specific History:
'
'Prog   Date        Description
'IK     25/05/2006  EP620, espom phase 1 (manual completion)
'GHun   14/11/2006  EP2_87 Changed CallExternalSystem to read progId and method from globals
'SR     18/01/2007  EP2_858 - GetNewNumbers is now through CallExternalSystem
'------------------------------------------------------------------------------------------
Public Function Request(ByVal vstrRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:
'   Main entry point for requests. Acts as a request broker
' pass:
'   vstrRequest
'       XML Input request specifying the data and operation
' return:
'   XML Response data stream containing results of operation
'------------------------------------------------------------------------------------------
On Error GoTo RequestVbErr
    
    Const cstrFunctionName As String = "Request"
    
    Dim xmlInDoc As FreeThreadedDOMDocument40
    Dim xmlOutDoc As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlOperationNode As IXMLDOMNode
    Dim xmlOperationNodeList As IXMLDOMNodeList
    Dim xmlAttribute As IXMLDOMAttribute

    m_objContext.SetComplete

    Set xmlOutDoc = New FreeThreadedDOMDocument40
    xmlOutDoc.validateOnParse = False
    xmlOutDoc.setProperty "NewParser", True
        
    Set xmlResponseElem = xmlOutDoc.createElement("RESPONSE")
    xmlOutDoc.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    Set xmlInDoc = xmlLoad(vstrRequest, cstrFunctionName)
        
    Set xmlRequestNode = xmlGetNode(xmlInDoc, "REQUEST")
   
    If Not xmlRequestNode.Attributes.getNamedItem("OPERATION") Is Nothing Then
        DoRequest xmlRequestNode, xmlResponseElem
    Else
        Set xmlOperationNodeList = xmlInDoc.selectNodes("REQUEST/OPERATION")
        For Each xmlOperationNode In xmlOperationNodeList
            For Each xmlAttribute In xmlRequestNode.Attributes
                xmlOperationNode.Attributes.setNamedItem xmlAttribute.cloneNode(True)
            Next
            DoRequest xmlOperationNode, xmlResponseElem
        Next
    End If

    Request = xmlOutDoc.xml

RequestExit:

    Set xmlInDoc = Nothing
    Set xmlOutDoc = Nothing
    Set xmlResponseElem = Nothing
    Set xmlRequestNode = Nothing
    Set xmlOperationNode = Nothing
    Set xmlOperationNodeList = Nothing
    Set xmlAttribute = Nothing
    
    Exit Function

RequestVbErr:
    
    Request = errCreateErrorResponse()
    
    GoTo RequestExit

End Function

Private Sub DoRequest( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   Determines which internal method to call
' pass:
'   vxmlRequestNode
'       XML Input request specifying the data and operation
'   vxmlResponseNode
'       XML Response to add the data to
' return:       N/A
'------------------------------------------------------------------------------------------
    On Error GoTo DoRequestExit
    
    Const cstrFunctionName As String = "DoRequest"
    
    Dim strOperation As String
    
    If vxmlRequestNode.nodeName = "REQUEST" Then
        strOperation = xmlGetMandatoryAttributeText(vxmlRequestNode, "OPERATION")
    Else
        strOperation = xmlGetMandatoryAttributeText(vxmlRequestNode, "NAME")
    End If
    
    strOperation = UCase(strOperation)
    
    Select Case strOperation
        
        Case "FINDCUSTOMER"
            'BMIDS00056 GHun Call BMids admin system
            'FindCustomer vxmlRequestNode, vxmlResponseNode
            CallExternalSystem vxmlRequestNode, vxmlResponseNode, "FindCustomer"
            'BMIDS00056 End
            
        Case "FINDBUSINESSFORCUSTOMER"
            'BMIDS00056 GHun Call BMids admin system
            'FindBusinessForCustomer vxmlRequestNode, vxmlResponseNode
            CallExternalSystem vxmlRequestNode, vxmlResponseNode, "FindBusinessForCustomer"
            'BMIDS00056 End
            
        Case "GETACCOUNTDETAILS"
            'BMIDS00006 GHun Call BMids admin system
            'GetAccountDetails vxmlRequestNode, vxmlResponseNode
            CallExternalSystem vxmlRequestNode, vxmlResponseNode, "GetAccountDetails"
            'BMIDS00006 End
            
        Case "GETACCOUNTSUMMARY"
            'BMIDS00056 GHun Call BMids admin system
            'GetAccountSummary vxmlRequestNode, vxmlResponseNode
            CallExternalSystem vxmlRequestNode, vxmlResponseNode, "GetAccountSummary"
            'BMIDS00056 End
            
        Case "GETCUSTOMERDETAILS"
            'BMIDS00056 GHun Call BMids admin system
            'GetCustomerDetails vxmlRequestNode, vxmlResponseNode
            CallExternalSystem vxmlRequestNode, vxmlResponseNode, "GetCustomerDetails"
            'BMIDS00056 End
            
        'SR 17/08/2004 : BMIDS815
        Case "GETACCOUNTREFRESH"
            CallExternalSystem vxmlRequestNode, vxmlResponseNode, "GetAccountRefresh"
        'BMIDS00056 End
        
        Case "GETNEWNUMBERS"
            CallExternalSystem vxmlRequestNode, vxmlResponseNode, "GetNewNumbers" 'SR 17/01/2007 : EP2_858
            'GetNewNumbers vxmlRequestNode, vxmlResponseNode
            
        Case "GETNEWNUMBERSWITHAUTOLOGON"
            GetNewNumbersWithAutoLogon vxmlRequestNode, vxmlResponseNode
                        
        Case "LOGOFFUSER"
            LogOffUser vxmlRequestNode, vxmlResponseNode
            
        Case "PROCESSCOMPLETIONSINTERFACE"
            ProcessCompletionsInterface vxmlRequestNode, vxmlResponseNode
            
        Case "VALIDATEUSERLOGON"
            ValidateUserLogon vxmlRequestNode, vxmlResponseNode
            
        'DM     02/04/02    SYS4350 Implement SaveThirdpartyDetails
        Case "SAVETHIRDPARTYDETAILS"
            SaveThirdPartyDetails vxmlRequestNode, vxmlResponseNode
           
        'MV - 21/06/2002 - BMIDS00077 - Core Upgrade
        'Start
        Case "GETUSERMORTGAGEADMINDETAILS"
            GetUserMortgageAdminDetails vxmlRequestNode, vxmlResponseNode

        
        Case "SAVEUSERMORTGAGEADMINDETAILS"
            SaveUserMortgageAdminDetails vxmlRequestNode, vxmlResponseNode

            
        Case "SAVEMORTGAGEPRODUCTDETAILS"
            SaveMortgageProductDetails vxmlRequestNode, vxmlResponseNode
        'End
        
        'BMIDS00005 Added GetCRSContactData
        Case "GETCRSCONTACTDATA"
            'GetCRSContactData vxmlRequestNode, vxmlResponseNode
            CallExternalSystem vxmlRequestNode, vxmlResponseNode, "GetCRSContactData"
        
        'BMIDS00005 Added UpdateCRSContactLog
        Case "UPDATECRSCONTACTLOG"
            'UpdateCRSContactLog vxmlRequestNode, vxmlResponseNode
            CallExternalSystem vxmlRequestNode, vxmlResponseNode, "UpdateCRSContactLog"
            
        'BMIDS00025 MDC 06/06/2002 - Add GetIntroducerData method
        Case "GETINTRODUCERDATA"
            'GetIntroducerData vxmlRequestNode, vxmlResponseNode
            CallExternalSystem vxmlRequestNode, vxmlResponseNode, "GetIntroducerData"
        'BMIDS00025 MDC 06/06/2002 - End
        
        'MO - 21/10/2002 - BMIDS00663 - FindIntroducer - Start
        Case "FINDINTRODUCER"
            CallExternalSystem vxmlRequestNode, vxmlResponseNode, "FindIntroducer"
        'MO - 21/10/2002 - BMIDS00663 - End
        
        'BMIDS00026 GHun 10/06/2002 - Add UpdateCRSCustomer method
        Case "UPDATECRSCUSTOMER"
            CallExternalSystem vxmlRequestNode, vxmlResponseNode, "UpdateCRSCustomer"
        'BMIDS00026 - End
        
        'BMIDS00006 MDC 16/07/2002 - CAWP1 BM054 Add GetAccountCustomers method
        Case "GETACCOUNTCUSTOMERS"
            CallExternalSystem vxmlRequestNode, vxmlResponseNode, "GetAccountCustomers"
        'BMIDS00006 MDC 16/07/2002 - End
        
        'BMIDS00336 MDC 20/08/2002 - CCWP1 BM062
        Case "EXPERIANCREDITCHECK"
            CallExternalSystem vxmlRequestNode, vxmlResponseNode, "ExperianCreditCheck"
        
        Case "EXPERIANBUREAU"
            CallExternalSystem vxmlRequestNode, vxmlResponseNode, "ExperianBureau"
        'BMIDS00336 MDC 20/08/2002 - End
        
        ' PSC 22/10/2002 BMIDS00676 - Start
        Case "RUNGENERICINTERFACE"
            RunGenericInterface vxmlRequestNode, vxmlResponseNode
        ' PSC 22/10/2002 BMIDS00676 - End
        
        Case Else
            errThrowError cstrFunctionName, oeNotImplemented, strOperation
    End Select
    
DoRequestExit:
    
    errCheckError cstrFunctionName, TypeName(Me)

End Sub


Private Sub ValidateUserLogon( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   Log on to the Administration system linked to Omiga.
'   Added for SYS2927 (Omiplus CC34).
' pass:
'   vxmlRequestNode
'       XML input request specifying the data and operation.
'   vxmlResponseNode
'       XML reponse to add the data to.
' return:
'   N/A
'------------------------------------------------------------------------------------------
On Error GoTo ValidateUserLogonExit
    
    Const cstrFunctionName As String = "ValidateUserLogon"
    
    Dim objODITransformer As Object
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim strTransformerResponse As String
    
    Set objODITransformer = CreateObject("ODITransformer.ODITransformerBO")
    strTransformerResponse = objODITransformer.Request(vxmlRequestNode.xml)
    Set xmlDoc = xmlLoad(strTransformerResponse, cstrFunctionName)
    errCheckXMLResponseNode xmlDoc.documentElement, vxmlResponseNode, True
    AttachResponseData vxmlResponseNode, xmlDoc.documentElement
    
ValidateUserLogonExit:
        
    Set objODITransformer = Nothing
    Set xmlDoc = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)

End Sub

Private Sub LogOffUser( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   Log off from the Administration system linked to Omiga.
'   Added for SYS2927 (Omiplus CC34).
' pass:
'   vxmlRequestNode
'       XML Input request specifying the data and operation
'   vxmlResponseNode
'       XML Reponse to add the data to
' return:
'   N/A
'------------------------------------------------------------------------------------------
On Error GoTo LogOffUserExit
    
    Const cstrFunctionName As String = "LogOffUser"

    Dim objODITransformer As Object
    Dim xmlDoc As FreeThreadedDOMDocument40
    Dim strTransformerResponse As String
    
    Set objODITransformer = CreateObject("ODITransformer.ODITransformerBO")
    strTransformerResponse = objODITransformer.Request(vxmlRequestNode.xml)
    Set xmlDoc = xmlLoad(strTransformerResponse, cstrFunctionName)
    errCheckXMLResponseNode xmlDoc.documentElement, vxmlResponseNode, True
    AttachResponseData vxmlResponseNode, xmlDoc.documentElement
    
LogOffUserExit:
        
    Set objODITransformer = Nothing
    Set xmlDoc = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)

End Sub

Private Sub GetNewNumbers(ByVal vxmlRequestNode As IXMLDOMNode, _
                          ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:  Get the related numbers from the Administration system linked to Omiga
' pass:         vxmlRequestNode  XML Input request specifying the data and operation
'               vxmlResponseNode XML Reponse to add the data to
' return:       N/A
'------------------------------------------------------------------------------------------
On Error GoTo GetNewNumbersExit

    Const cstrFunctionName As String = "GetNewNumbers"

    Dim strNumberRequired As String, strTypeRequired As String, strOtherSysNumber As String
    Dim xmlNumberRequired As IXMLDOMElement, xmlOtherSystemNumber As IXMLDOMElement
    Dim xmlDoc As New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    Dim xmlTransformerRequest As New FreeThreadedDOMDocument40
    xmlTransformerRequest.validateOnParse = False
    xmlTransformerRequest.setProperty "NewParser", True
    Dim xmlTransformerResponse As New FreeThreadedDOMDocument40
    xmlTransformerResponse.validateOnParse = False
    xmlTransformerResponse.setProperty "NewParser", True
    Dim xmlTransformerNode As IXMLDOMNode
    Dim objODITransformer As Object
    Dim strTransformerResponse As String

    If (GetGlobalParamBoolean("FindLegacyCustomer") = "1") Then

        Set xmlNumberRequired = xmlGetMandatoryNode(vxmlRequestNode, ".//NUMBERREQUEST")
        strNumberRequired = xmlGetAttributeText(xmlNumberRequired, "NUMBERREQUIRED")

        Set xmlTransformerNode = xmlTransformerRequest.createElement("REQUEST")
        xmlTransformerRequest.appendChild xmlTransformerNode
        xmlCopyAttribute vxmlRequestNode, xmlTransformerNode, "ADMINSYSTEMSTATE"

        If strNumberRequired <> "" Then
            Set xmlOtherSystemNumber = xmlDoc.createElement("OTHERSYSTEMNUMBER")
            vxmlResponseNode.appendChild xmlOtherSystemNumber

            strTypeRequired = xmlGetAttributeText(xmlNumberRequired, "NUMBERTYPEREQUIRED")

            If strTypeRequired = "A" Then
                xmlSetAttributeValue xmlTransformerNode, "OPERATION", "GetNewAccountNumber"
                Set objODITransformer = CreateObject("ODITransformer.ODITransformerBO")
                strTransformerResponse = objODITransformer.Request(xmlTransformerNode.xml)
                xmlTransformerResponse.loadXML (strTransformerResponse)
                errCheckXMLResponseNode xmlTransformerResponse.selectSingleNode("RESPONSE"), vxmlResponseNode, True


                strOtherSysNumber = xmlTransformerResponse.selectSingleNode("RESPONSE/NUMBERRESPONSE").Attributes.getNamedItem("OTHERSYSTEMNUMBER").Text

                'strOtherSysNumber = "OPT" + "A" + strNumberRequired

                xmlSetAttributeValue xmlOtherSystemNumber, "OTHERSYSTEMNUMBER", strOtherSysNumber
                xmlSetAttributeValue xmlOtherSystemNumber, "OMIGANUMBER", strNumberRequired
                xmlSetAttributeValue xmlOtherSystemNumber, "NUMBERTYPE", strTypeRequired

            ElseIf strTypeRequired = "C" Then
                xmlSetAttributeValue xmlTransformerNode, "OPERATION", "GetNewCustomerNumber"
                Set objODITransformer = CreateObject("ODITransformer.ODITransformerBO")
                strTransformerResponse = objODITransformer.Request(xmlTransformerNode.xml)
                xmlTransformerResponse.loadXML (strTransformerResponse)
                errCheckXMLResponseNode xmlTransformerResponse.selectSingleNode("RESPONSE"), vxmlResponseNode, True

                
                strOtherSysNumber = xmlTransformerResponse.selectSingleNode("RESPONSE/NUMBERRESPONSE").Attributes.getNamedItem("OTHERSYSTEMNUMBER").Text

                'strOtherSysNumber = "OPT" + "C" + strNumberRequired

                xmlSetAttributeValue xmlOtherSystemNumber, "OTHERSYSTEMNUMBER", strOtherSysNumber
                xmlSetAttributeValue xmlOtherSystemNumber, "NUMBERREQUIRED", strNumberRequired
                xmlSetAttributeValue xmlOtherSystemNumber, "NUMBERTYPE", strTypeRequired
            Else
                errThrowError "GetNewNumbers", 4908
            End If
        Else
            errThrowError "GetNewNumbers", 4907
        End If
    End If

GetNewNumbersExit:

    Set xmlNumberRequired = Nothing
    Set xmlOtherSystemNumber = Nothing
    Set xmlDoc = Nothing

    errCheckError cstrFunctionName, TypeName(Me)

End Sub

Private Sub GetNewNumbersWithAutoLogon(ByVal vxmlRequestNode As IXMLDOMNode, _
                          ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:  Get the related numbers from the Administration system linked to Omiga
' pass:         vxmlRequestNode  XML Input request specifying the data and operation
'               vxmlResponseNode XML Reponse to add the data to
' return:       N/A
'------------------------------------------------------------------------------------------
On Error GoTo GetNewNumbersWithAutoLogonExit

    Const cstrFunctionName As String = "GetNewNumbersWithAutoLogon"

    Dim strNumberRequired As String, strTypeRequired As String, strOtherSysNumber As String
    Dim xmlNumberRequired As IXMLDOMElement, xmlOtherSystemNumber As IXMLDOMElement
    Dim xmlDoc As New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    Dim xmlTransformerRequest As New FreeThreadedDOMDocument40
    xmlTransformerRequest.validateOnParse = False
    xmlTransformerRequest.setProperty "NewParser", True
    Dim xmlTransformerResponse As New FreeThreadedDOMDocument40
    xmlTransformerResponse.validateOnParse = False
    xmlTransformerResponse.setProperty "NewParser", True
    Dim xmlTransformerNode As IXMLDOMNode
    Dim objODITransformer As Object
    Dim strTransformerResponse As String
    Dim strLogRequest, strLogResponse, strAdminSystemState As String
    Dim xmlLogDoc As New FreeThreadedDOMDocument40
    xmlLogDoc.validateOnParse = False
    xmlLogDoc.setProperty "NewParser", True
    Dim xmlLogRequestNode, xmlLogResponseNode As IXMLDOMNode

    If (GetGlobalParamBoolean("FindLegacyCustomer") = "1") Then
    
        ' Login to the admin system
                        
        strLogRequest = "<REQUEST OPERATION=""ValidateUserLogon"" CHANNELID=""A1""><USER ENVIRONMENT = """ + GetGlobalParamString("ODIAutoLogonUserEnv") + """ HOST = """ + GetGlobalParamString("ODIAutoLogonHost") + """ USERNAME = """ _
                        & GetGlobalParamString("ODIAutoLogonUserName") & """ PASSWORDVALUE = """ _
                        & GetGlobalParamString("ODIAutoLogonPassword") & """/><ODIINITIALISATION ODIENVIRONMENT = """ + GetGlobalParamString("ODIAutoLogonODIEnv") + """/></REQUEST>"
        Set objODITransformer = CreateObject("ODITransformer.ODITransformerBO")
        strLogResponse = objODITransformer.Request(strLogRequest)
        xmlLogDoc.loadXML (strLogResponse)
        errCheckXMLResponseNode xmlLogDoc.selectSingleNode("RESPONSE"), vxmlResponseNode, True

        strAdminSystemState = xmlLogDoc.selectSingleNode("RESPONSE/ADMINSYSTEMSTATE").xml
        
        Set xmlNumberRequired = xmlGetMandatoryNode(vxmlRequestNode, ".//NUMBERREQUEST")
        strNumberRequired = xmlGetAttributeText(xmlNumberRequired, "NUMBERREQUIRED")

        Set xmlTransformerNode = xmlTransformerRequest.createElement("REQUEST")
        xmlTransformerRequest.appendChild xmlTransformerNode
        
        xmlSetAttributeValue xmlTransformerNode, "ADMINSYSTEMSTATE", strAdminSystemState

        If strNumberRequired <> "" Then
            Set xmlOtherSystemNumber = xmlDoc.createElement("OTHERSYSTEMNUMBER")
            vxmlResponseNode.appendChild xmlOtherSystemNumber

            strTypeRequired = xmlGetAttributeText(xmlNumberRequired, "NUMBERTYPEREQUIRED")

            If strTypeRequired = "A" Then
                xmlSetAttributeValue xmlTransformerNode, "OPERATION", "GetNewAccountNumber"
                strTransformerResponse = objODITransformer.Request(xmlTransformerNode.xml)
                xmlTransformerResponse.loadXML (strTransformerResponse)
                errCheckXMLResponseNode xmlTransformerResponse.selectSingleNode("RESPONSE")

                strOtherSysNumber = xmlTransformerResponse.selectSingleNode("RESPONSE/NUMBERRESPONSE").Attributes.getNamedItem("OTHERSYSTEMNUMBER").Text
                xmlSetAttributeValue xmlOtherSystemNumber, "OTHERSYSTEMNUMBER", strOtherSysNumber
                xmlSetAttributeValue xmlOtherSystemNumber, "OMIGANUMBER", strNumberRequired
                xmlSetAttributeValue xmlOtherSystemNumber, "NUMBERTYPE", strTypeRequired

            ElseIf strTypeRequired = "C" Then
                xmlSetAttributeValue xmlTransformerNode, "OPERATION", "GetNewCustomerNumber"
                strTransformerResponse = objODITransformer.Request(xmlTransformerNode.xml)
                xmlTransformerResponse.loadXML (strTransformerResponse)
                errCheckXMLResponseNode xmlTransformerResponse.selectSingleNode("RESPONSE")

                
                strOtherSysNumber = xmlTransformerResponse.selectSingleNode("RESPONSE/NUMBERRESPONSE").Attributes.getNamedItem("OTHERSYSTEMNUMBER").Text

                'strOtherSysNumber = "OPT" + "C" + strNumberRequired

                xmlSetAttributeValue xmlOtherSystemNumber, "OTHERSYSTEMNUMBER", strOtherSysNumber
                xmlSetAttributeValue xmlOtherSystemNumber, "NUMBERREQUIRED", strNumberRequired
                xmlSetAttributeValue xmlOtherSystemNumber, "NUMBERTYPE", strTypeRequired
            Else
                errThrowError "GetNewNumbersWithAutoLogon", 4908
            End If
            
            'Log off the admin system
            Set xmlLogDoc = New FreeThreadedDOMDocument40
            xmlLogDoc.validateOnParse = False
            xmlLogDoc.setProperty "NewParser", True
            xmlLogDoc.loadXML ("<REQUEST OPERATION=""LogOffUser""></REQUEST>")
            xmlSetAttributeValue xmlLogDoc.selectSingleNode("REQUEST"), "ADMINSYSTEMSTATE", strAdminSystemState
            Set objODITransformer = CreateObject("ODITransformer.ODITransformerBO")
            strLogResponse = objODITransformer.Request(xmlLogDoc.xml)
            Set xmlLogDoc = New FreeThreadedDOMDocument40
            xmlLogDoc.validateOnParse = False
            xmlLogDoc.setProperty "NewParser", True
            xmlLogDoc.loadXML (strLogResponse)
            errCheckXMLResponseNode xmlLogDoc.selectSingleNode("RESPONSE"), vxmlResponseNode, True
            
        Else
            errThrowError "GetNewNumbersWithAutoLogon", 4907
        End If
    End If

GetNewNumbersWithAutoLogonExit:

    Set xmlNumberRequired = Nothing
    Set xmlOtherSystemNumber = Nothing
    Set xmlDoc = Nothing

    errCheckError cstrFunctionName, TypeName(Me)

End Sub

Private Function FindCustomer(ByVal vxmlRequestNode As IXMLDOMNode, _
                              ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:  Finds a list of customers matching the criteria
' pass:         vxmlRequestNode  XML Input request specifying the data and operation
'               vxmlResponseNode XML Reponse to add the data to
' return:       N/A
'------------------------------------------------------------------------------------------
On Error GoTo FindCustomerExit

    Const cstrFunctionName As String = "FindCustomer"
    
    Dim xmlCustomerList As IXMLDOMElement
    Dim xmlCustomer As IXMLDOMElement
    Dim xmlTransformerResponse As New FreeThreadedDOMDocument40
    xmlTransformerResponse.validateOnParse = False
    xmlTransformerResponse.setProperty "NewParser", True
    Dim strTransformerResponse As String
    Dim objODITransformer As Object
    Dim intCount As Integer
    Dim strCount As String
    Dim strChannelId As String
    
    strChannelId = xmlGetMandatoryAttributeText(vxmlRequestNode, "CHANNELID")
    
    'DS 12/11/2001 OmiPlus29 Add ADMINSYSTEMSTATE to response
    vxmlResponseNode.ownerDocument.createElement ("ADMINSYSTEMSTATE")
    
    vxmlRequestNode.Attributes.getNamedItem("OPERATION").Text = "FindCustomerList"
    
    Set objODITransformer = CreateObject("ODITransformer.ODITransformerBO")
    strTransformerResponse = objODITransformer.Request(vxmlRequestNode.xml)
    
    xmlTransformerResponse.loadXML (strTransformerResponse)
    errCheckXMLResponseNode xmlTransformerResponse.selectSingleNode("RESPONSE"), vxmlResponseNode, True
    
    Dim xmlChildNode As IXMLDOMNode
    For Each xmlChildNode In xmlTransformerResponse.selectSingleNode("RESPONSE").childNodes
        vxmlResponseNode.appendChild xmlChildNode.cloneNode(True)
    Next
    
FindCustomerExit:
    
    Set xmlCustomerList = Nothing
    Set xmlCustomer = Nothing
   
    errCheckError cstrFunctionName, TypeName(Me)

End Function

Private Function GetCustomerDetails(ByVal vxmlRequestNode As IXMLDOMNode, _
                                    ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:  Gets the details of the customer passed in
' pass:         vxmlRequestNode  XML Input request specifying the data and operation
'               vxmlResponseNode XML Reponse to add the data to
' return:       N/A
'------------------------------------------------------------------------------------------
On Error GoTo GetCustomerDetailsExit

    Const cstrFunctionName As String = "GetCustomerDetails"
                                    
    Dim xmlTransformerResponse As New FreeThreadedDOMDocument40
    xmlTransformerResponse.validateOnParse = False
    xmlTransformerResponse.setProperty "NewParser", True
    Dim xmlChildNode As IXMLDOMNode
    Dim objODITransformer As Object
    Dim strTransformerResponse As String
    
    Set objODITransformer = CreateObject("ODITransformer.ODITransformerBO")
    strTransformerResponse = objODITransformer.Request(vxmlRequestNode.xml)
    xmlTransformerResponse.loadXML (strTransformerResponse)
    errCheckXMLResponseNode xmlTransformerResponse.selectSingleNode("RESPONSE"), vxmlResponseNode, True
    
    For Each xmlChildNode In xmlTransformerResponse.selectSingleNode("RESPONSE").childNodes
        vxmlResponseNode.appendChild xmlChildNode.cloneNode(True)
    Next
    
GetCustomerDetailsExit:
    Set xmlTransformerResponse = Nothing
    Set xmlChildNode = Nothing
    Set objODITransformer = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)

End Function

Private Function FindBusinessForCustomer(ByVal vxmlRequestNode As IXMLDOMNode, _
                                         ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:  Finds a list of the business connected to the customer passed in
' pass:         vxmlRequestNode  XML Input request specifying the data and operation
'               vxmlResponseNode XML Reponse to add the data to
' return:       N/A
'------------------------------------------------------------------------------------------
On Error GoTo FindBusinessForCustomerError

    Const cstrFunctionName As String = "FindBusinessForCustomer"

    Dim xmlTransformerResponse As New FreeThreadedDOMDocument40
    xmlTransformerResponse.validateOnParse = False
    xmlTransformerResponse.setProperty "NewParser", True
    Dim xmlChildNode As IXMLDOMNode
    Dim objODITransformer As Object
    Dim strTransformerResponse As String
    
    Set objODITransformer = CreateObject("ODITransformer.ODITransformerBO")
    strTransformerResponse = objODITransformer.Request(vxmlRequestNode.xml)
    xmlTransformerResponse.loadXML (strTransformerResponse)
    errCheckXMLResponseNode xmlTransformerResponse.selectSingleNode("RESPONSE"), vxmlResponseNode, True
    
    For Each xmlChildNode In xmlTransformerResponse.selectSingleNode("RESPONSE").childNodes
        vxmlResponseNode.appendChild xmlChildNode.cloneNode(True)
    Next
            
FindBusinessForCustomerExit:
         
    errCheckError cstrFunctionName, TypeName(Me)
    
FindBusinessForCustomerError:

    errCheckError cstrFunctionName, TypeName(Me)

End Function

Private Function GetAccountDetails(ByVal vxmlRequestNode As IXMLDOMNode, _
                                   ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:  Gets the account details for the customer passed in
' pass:         vxmlRequestNode  XML Input request specifying the data and operation
'               vxmlResponseNode XML Reponse to add the data to
' return:       N/A
'------------------------------------------------------------------------------------------
On Error GoTo GetAccountDetailsExit

    Const cstrFunctionName As String = "GetAccountDetails"
    
    Dim xmlTransformerResponse As New FreeThreadedDOMDocument40
    xmlTransformerResponse.validateOnParse = False
    xmlTransformerResponse.setProperty "NewParser", True
    Dim xmlChildNode As IXMLDOMNode
    Dim objODITransformer As Object
    Dim strTransformerResponse As String
    
    vxmlRequestNode.Attributes.getNamedItem("OPERATION").Text = "FindAccountDetails"
    
    Set objODITransformer = CreateObject("ODITransformer.ODITransformerBO")
    strTransformerResponse = objODITransformer.Request(vxmlRequestNode.xml)
    xmlTransformerResponse.loadXML (strTransformerResponse)
    errCheckXMLResponseNode xmlTransformerResponse.selectSingleNode("RESPONSE"), vxmlResponseNode, True
    
    For Each xmlChildNode In xmlTransformerResponse.selectSingleNode("RESPONSE").childNodes
        vxmlResponseNode.appendChild xmlChildNode.cloneNode(True)
    Next
    
'
'
'    Dim xmlAccountList As IXMLDOMElement
'    Dim xmlCustomerList As IXMLDOMElement
'    Dim xmlCustomer As IXMLDOMElement
'    Dim xmlPhoneList As IXMLDOMElement
'    Dim xmlLoanList As IXMLDOMElement
'    Dim xmlAccount As IXMLDOMElement
'    Dim xmlProperty As IXMLDOMElement
'    Dim xmlElement As IXMLDOMElement
'
'    Dim strChannelId As String
'
'    strChannelId = xmlGetMandatoryAttributeText(vxmlRequestNode, "CHANNELID")
'
'    Set xmlAccountList = vxmlResponseNode.ownerDocument.createElement("MORTGAGEACCOUNTLIST")
'    vxmlResponseNode.appendChild xmlAccountList
'
'    Dim lngCount As Long
'    Dim strCount As String
'
'    For lngCount = 1 To 2
'
'        Set xmlAccount = vxmlResponseNode.ownerDocument.createElement("MORTGAGEACCOUNT")
'        xmlAccountList.appendChild xmlAccount
'
'        strCount = Format(lngCount, "00")
'
'        xmlAccount.setAttribute "ACCOUNTNUMBER", "AC00000000" & strCount
'        xmlAccount.setAttribute "SECONDCHARGEINDICATOR", "1"
'        xmlAccount.setAttribute "INDEMNITYCOMPANYNAME", "The MIG Company"
'        xmlAccount.setAttribute "INDEMNITYMORTGAGEAMOUNT", CStr(5000 * lngCount)
'        xmlAccount.setAttribute "INDEMNITYAMOUNT", CStr(1000 * lngCount)
'        xmlAccount.setAttribute "LENDERCODE", "2(BRL)"
'        xmlAccount.setAttribute "CREATIONDATE", strCount & "/" & strCount & "/19" & CStr(60 + lngCount)
'
'        Set xmlCustomerList = vxmlResponseNode.ownerDocument.createElement("CUSTOMERLIST")
'        xmlAccount.appendChild xmlCustomerList
'
'        Set xmlCustomer = vxmlResponseNode.ownerDocument.createElement("CUSTOMER")
'        xmlCustomerList.appendChild xmlCustomer
'
'        xmlCustomer.setAttribute "CUSTOMERNUMBER", "ADMIN0000001"
'        xmlCustomer.setAttribute "SURNAME", "Jones"
'        xmlCustomer.setAttribute "FIRSTFORENAME", "Frederick"
'        xmlCustomer.setAttribute "SECONDFORENAME", "Jeffrey"
'        xmlCustomer.setAttribute "OTHERFORENAME", "Arnold"
'        xmlCustomer.setAttribute "TITLE", "1"
'        xmlCustomer.setAttribute "DATEOFBIRTH", "12/05/1968"
'        xmlCustomer.setAttribute "AGE", DateDiff("yyyy", "12/05/1968", Date)
'        xmlCustomer.setAttribute "GENDER", "1"
'        xmlCustomer.setAttribute "MOTHERSMAIDENNAME", "Smith"
'        xmlCustomer.setAttribute "CONTACTEMAILADDRESS", "fred.jones@company.com"
'        xmlCustomer.setAttribute "MARITALSTATUS", "1"
'        xmlCustomer.setAttribute "NATIONALINSURANCENUMBER", "AB563143D"
'        xmlCustomer.setAttribute "MEMBEROFSTAFF", "0"
'        xmlCustomer.setAttribute "CUSTOMERORDER", "2"
'        xmlCustomer.setAttribute "CUSTOMERROLETYPE", "1"
'        xmlCustomer.setAttribute "CHANNELID", strChannelId
'        xmlCustomer.setAttribute "OTHERSYSTEMTYPE", "1"
'
'
'        Set xmlElement = vxmlResponseNode.ownerDocument.createElement("CURRENTADDRESS")
'        xmlCustomer.appendChild xmlElement
'
'        xmlElement.setAttribute "BUILDINGORHOUSENAME", "The Larches"
'        xmlElement.setAttribute "BUILDINGORHOUSENUMBER", "12"
'        xmlElement.setAttribute "FLATNUMBER", "12"
'        xmlElement.setAttribute "STREET", "Pond Lane"
'        xmlElement.setAttribute "DISTRICT", "Blockley"
'        xmlElement.setAttribute "TOWN", "Chipping Campden"
'        xmlElement.setAttribute "COUNTY", "Gloucestershire"
'        xmlElement.setAttribute "COUNTRY", "1"
'        xmlElement.setAttribute "POSTCODE", "GL12 3SH"
'        xmlElement.setAttribute "DELIVERYPOINTSUFFIX", "X"
'        xmlElement.setAttribute "MAILSORTCODE", "2"
'
'        Set xmlElement = vxmlResponseNode.ownerDocument.createElement("CORRESPONDENCEADDRESS")
'        xmlCustomer.appendChild xmlElement
'
'        xmlElement.setAttribute "BUILDINGORHOUSENAME", "The Place"
'        xmlElement.setAttribute "BUILDINGORHOUSENUMBER", "345"
'        xmlElement.setAttribute "FLATNUMBER", ""
'        xmlElement.setAttribute "STREET", "Brown Street"
'        xmlElement.setAttribute "DISTRICT", "Montpellier"
'        xmlElement.setAttribute "TOWN", "Cheltenham"
'        xmlElement.setAttribute "COUNTY", "Gloucestershire"
'        xmlElement.setAttribute "COUNTRY", "1"
'        xmlElement.setAttribute "POSTCODE", "GL50 3SH"
'        xmlElement.setAttribute "DELIVERYPOINTSUFFIX", "Y"
'        xmlElement.setAttribute "MAILSORTCODE", "3"
'
'        Set xmlPhoneList = vxmlResponseNode.ownerDocument.createElement("TELEPHONENUMBERLIST")
'        xmlCustomer.appendChild xmlPhoneList
'
'        Set xmlElement = vxmlResponseNode.ownerDocument.createElement("TELEPHONENUMBERDETAILS")
'        xmlPhoneList.appendChild xmlElement
'
'        xmlElement.setAttribute "USAGE", "1"
'        xmlElement.setAttribute "TELEPHONENUMBER", "01242 123412"
'        xmlElement.setAttribute "EXTENSIONNUMBER", ""
'        xmlElement.setAttribute "CONTACTTIME", "Anytime"
'        xmlElement.setAttribute "PREFERREDMETHODOFCONTACT", "1"
'
'        Set xmlElement = vxmlResponseNode.ownerDocument.createElement("TELEPHONENUMBERDETAILS")
'        xmlPhoneList.appendChild xmlElement
'
'        xmlElement.setAttribute "USAGE", "2"
'        xmlElement.setAttribute "TELEPHONENUMBER", "01242 123475"
'        xmlElement.setAttribute "EXTENSIONNUMBER", "1234"
'        xmlElement.setAttribute "CONTACTTIME", "Anytime"
'        xmlElement.setAttribute "PREFERREDMETHODOFCONTACT", "0"
'
'
'        Set xmlCustomer = vxmlResponseNode.ownerDocument.createElement("CUSTOMER")
'        xmlCustomerList.appendChild xmlCustomer
'
'        xmlCustomer.setAttribute "CUSTOMERNUMBER", "ADMIN0000100"
'        xmlCustomer.setAttribute "SURNAME", "Jones"
'        xmlCustomer.setAttribute "FIRSTFORENAME", "Mary"
'        xmlCustomer.setAttribute "SECONDFORENAME", "Jane"
'        xmlCustomer.setAttribute "OTHERFORENAME", "Doris"
'        xmlCustomer.setAttribute "TITLE", "2"
'        xmlCustomer.setAttribute "DATEOFBIRTH", "10/01/1970"
'        xmlCustomer.setAttribute "AGE", DateDiff("yyyy", "10/01/1970", Date)
'        xmlCustomer.setAttribute "GENDER", "2"
'        xmlCustomer.setAttribute "MOTHERSMAIDENNAME", "Brown"
'        xmlCustomer.setAttribute "CONTACTEMAILADDRESS", "mary.jones@somothercompany.com"
'        xmlCustomer.setAttribute "MARITALSTATUS", "1"
'        xmlCustomer.setAttribute "NATIONALINSURANCENUMBER", "NL563143D"
'        xmlCustomer.setAttribute "MEMBEROFSTAFF", "0"
'        xmlCustomer.setAttribute "CUSTOMERORDER", "1"
'        xmlCustomer.setAttribute "CUSTOMERROLETYPE", "1"
'        xmlCustomer.setAttribute "CHANNELID", strChannelId
'        xmlCustomer.setAttribute "OTHERSYSTEMTYPE", "1"
'
'
'        Set xmlElement = vxmlResponseNode.ownerDocument.createElement("CURRENTADDRESS")
'        xmlCustomer.appendChild xmlElement
'
'        xmlElement.setAttribute "BUILDINGORHOUSENAME", "The Larches"
'        xmlElement.setAttribute "BUILDINGORHOUSENUMBER", "12"
'        xmlElement.setAttribute "FLATNUMBER", "12"
'        xmlElement.setAttribute "STREET", "Pond Lane"
'        xmlElement.setAttribute "DISTRICT", "Blockley"
'        xmlElement.setAttribute "TOWN", "Chipping Campden"
'        xmlElement.setAttribute "COUNTY", "Gloucestershire"
'        xmlElement.setAttribute "COUNTRY", "1"
'        xmlElement.setAttribute "POSTCODE", "GL12 3SH"
'        xmlElement.setAttribute "DELIVERYPOINTSUFFIX", "X"
'        xmlElement.setAttribute "MAILSORTCODE", "2"
'
'        Set xmlElement = vxmlResponseNode.ownerDocument.createElement("CORRESPONDENCEADDRESS")
'        xmlCustomer.appendChild xmlElement
'
'        xmlElement.setAttribute "BUILDINGORHOUSENAME", "The Place"
'        xmlElement.setAttribute "BUILDINGORHOUSENUMBER", "345"
'        xmlElement.setAttribute "FLATNUMBER", ""
'        xmlElement.setAttribute "STREET", "Brown Street"
'        xmlElement.setAttribute "DISTRICT", "Montpellier"
'        xmlElement.setAttribute "TOWN", "Cheltenham"
'        xmlElement.setAttribute "COUNTY", "Gloucestershire"
'        xmlElement.setAttribute "COUNTRY", "1"
'        xmlElement.setAttribute "POSTCODE", "GL50 3SH"
'        xmlElement.setAttribute "DELIVERYPOINTSUFFIX", "Y"
'        xmlElement.setAttribute "MAILSORTCODE", "3"
'
'        Set xmlPhoneList = vxmlResponseNode.ownerDocument.createElement("TELEPHONENUMBERLIST")
'        xmlCustomer.appendChild xmlPhoneList
'
'        Set xmlElement = vxmlResponseNode.ownerDocument.createElement("TELEPHONENUMBERDETAILS")
'        xmlPhoneList.appendChild xmlElement
'
'        xmlElement.setAttribute "USAGE", "1"
'        xmlElement.setAttribute "TELEPHONENUMBER", "01242 123412"
'        xmlElement.setAttribute "EXTENSIONNUMBER", ""
'        xmlElement.setAttribute "CONTACTTIME", "Anytime"
'        xmlElement.setAttribute "PREFERREDMETHODOFCONTACT", "1"
'
'        Set xmlElement = vxmlResponseNode.ownerDocument.createElement("TELEPHONENUMBERDETAILS")
'        xmlPhoneList.appendChild xmlElement
'
'        xmlElement.setAttribute "USAGE", "2"
'        xmlElement.setAttribute "TELEPHONENUMBER", "01242 1236868"
'        xmlElement.setAttribute "EXTENSIONNUMBER", "1234"
'        xmlElement.setAttribute "CONTACTTIME", "Anytime"
'        xmlElement.setAttribute "PREFERREDMETHODOFCONTACT", "0"
'
'        Set xmlProperty = vxmlResponseNode.ownerDocument.createElement("PROPERTY")
'        xmlAccount.appendChild xmlProperty
'
'        xmlProperty.setAttribute "PURCHASEPRICE", CStr(lngCount * 80000)
'        xmlProperty.setAttribute "VALUATIONDATE", "01/05/1990"
'        xmlProperty.setAttribute "VALUATIONAMOUNT", CStr(lngCount * 85000)
'        xmlProperty.setAttribute "VALUERID", "1234"
'        xmlProperty.setAttribute "PROPERTYDESCRIPTION", "1"
'        xmlProperty.setAttribute "YEARBUILT", "1954"
'        xmlProperty.setAttribute "PROPERTYTENURE", "1"
'        xmlProperty.setAttribute "BUILDINGSSUMINSURED", CStr(lngCount * 81000)
'        xmlProperty.setAttribute "REINSTATEMENTAMOUNT", CStr(lngCount * 86000)
'        xmlProperty.setAttribute "HOMEINSURANCETYPE", "1"
'
'        Set xmlElement = vxmlResponseNode.ownerDocument.createElement("PROPERTYADDRESS")
'        xmlProperty.appendChild xmlElement
'
'        xmlElement.setAttribute "BUILDINGORHOUSENAME", "The Place"
'        xmlElement.setAttribute "BUILDINGORHOUSENUMBER", "345"
'        xmlElement.setAttribute "FLATNUMBER", ""
'        xmlElement.setAttribute "STREET", "Bath Street"
'        xmlElement.setAttribute "DISTRICT", "Bishops's Cleeve"
'        xmlElement.setAttribute "TOWN", "Cheltenham"
'        xmlElement.setAttribute "COUNTY", "Gloucestershire"
'        xmlElement.setAttribute "COUNTRY", "1"
'        xmlElement.setAttribute "POSTCODE", "GL57 9SH"
'        xmlElement.setAttribute "DELIVERYPOINTSUFFIX", "Z"
'        xmlElement.setAttribute "MAILSORTCODE", "5"
'
'        Set xmlLoanList = vxmlResponseNode.ownerDocument.createElement("MORTGAGELOANLIST")
'        xmlAccount.appendChild xmlLoanList
'
'        Set xmlElement = vxmlResponseNode.ownerDocument.createElement("MORTGAGELOAN")
'        xmlLoanList.appendChild xmlElement
'
'        xmlElement.setAttribute "ACCOUNTNUMBER", "AC00000000" & strCount
'        xmlElement.setAttribute "LOANACCOUNTNUMBER", "AC00000000" & strCount & "/1"
'        xmlElement.setAttribute "MORTGAGEPRODUCTCODE", "150"
'        xmlElement.setAttribute "MORTGAGEPRODUCTDESCRIPTION", "BRL Fixed 3.99% to 01/07/2000"
'        xmlElement.setAttribute "INTERESTRATE", "3.99"
'        xmlElement.setAttribute "INTERESTRATETYPE", "F"
'        xmlElement.setAttribute "PURPOSEOFLOAN", "1"
'        xmlElement.setAttribute "OUTSTANDINGBALANCE", "20000"
'        xmlElement.setAttribute "REPAYMENTTYPE", "1"
'        xmlElement.setAttribute "ORIGINALLOANAMOUNT", "50000"
'        xmlElement.setAttribute "MONTHLYREPAYMENT", "200"
'        xmlElement.setAttribute "ORIGINALTERMYEARS", "25"
'        xmlElement.setAttribute "ORIGINALTERMMONTHS", "0"
'        xmlElement.setAttribute "STARTDATE", "01/03/1980"
'        xmlElement.setAttribute "REDEMPTIONSTATUS", "3"
'        xmlElement.setAttribute "REDEMPTIONDATE", "01/04/2001"
'
'        Set xmlElement = vxmlResponseNode.ownerDocument.createElement("MORTGAGELOAN")
'        xmlLoanList.appendChild xmlElement
'
'        xmlElement.setAttribute "ACCOUNTNUMBER", "AC00000000" & strCount
'        xmlElement.setAttribute "LOANACCOUNTNUMBER", "AC00000000" & strCount & "/2"
'        xmlElement.setAttribute "MORTGAGEPRODUCTCODE", "605"
'        xmlElement.setAttribute "MORTGAGEPRODUCTDESCRIPTION", "BRL Base Rate Large / 100% AirMiles "
'        xmlElement.setAttribute "INTERESTRATE", "7.67"
'        xmlElement.setAttribute "INTERESTRATETYPE", "B"
'        xmlElement.setAttribute "PURPOSEOFLOAN", "1"
'        xmlElement.setAttribute "OUTSTANDINGBALANCE", "40000"
'        xmlElement.setAttribute "REPAYMENTTYPE", "1"
'        xmlElement.setAttribute "ORIGINALLOANAMOUNT", "60000"
'        xmlElement.setAttribute "MONTHLYREPAYMENT", "250"
'        xmlElement.setAttribute "ORIGINALTERMYEARS", "25"
'        xmlElement.setAttribute "ORIGINALTERMMONTHS", "0"
'        xmlElement.setAttribute "STARTDATE", "01/03/1980"
'        xmlElement.setAttribute "REDEMPTIONSTATUS", "3"
'        xmlElement.setAttribute "REDEMPTIONDATE", "01/04/2001"
'
'        Set xmlElement = vxmlResponseNode.ownerDocument.createElement("ARREARSHISTORY")
'        xmlAccount.appendChild xmlElement
'
'        xmlElement.setAttribute "MAXIMUMBALANCE", strCount & "00"
'        xmlElement.setAttribute "MAXIMUMNUMBEROFMONTHS", strCount & "0"
'        xmlElement.setAttribute "DATECLEARED", "01/09/1990"
'        xmlElement.setAttribute "DESCRIPTIONOFLOAN", "1"
'
'    Next
'
GetAccountDetailsExit:
        
'    Set xmlAccountList = Nothing
'    Set xmlCustomerList = Nothing
'    Set xmlCustomer = Nothing
'    Set xmlPhoneList = Nothing
'    Set xmlLoanList = Nothing
'    Set xmlAccount = Nothing
'    Set xmlProperty = Nothing
'    Set xmlElement = Nothing

   
    errCheckError cstrFunctionName, TypeName(Me)

End Function

Private Sub ProcessCompletionsInterface( _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:
'   (usually) Generates completions request and passes it to external mortgage admin system
'   EPSOM (phase 1) completes batch only
' pass:
'   vxmlRequestNode  XML Input request specifying the data and operation
'   vxmlResponseNode XML Reponse to add the data to
' return:       N/A
' history:
'   MC  27/04/01    Created
'   RF  13/12/01    SYS3489 (Omiplus CC32) Call omToCBA (via OmiQueue)
'   IK  25/05/2006  EP620, espom phase 1 (manual completion)
'------------------------------------------------------------------------------------------
On Error GoTo ProcessCompletionsInterface_Exit

    Const strFunctionName As String = "ProcessCompletionsInterface"
    
    Dim omObj As Object
    Dim xmlReqDoc As FreeThreadedDOMDocument40
    Dim xmlRespDoc As FreeThreadedDOMDocument40
    Dim xmlElem As IXMLDOMElement
    Dim strResponse As String
    
    'ik_debug
    'vxmlRequestNode.ownerDocument.save "c:\omiga4Trace\ProcessCompletionsInterfaceRequest.xml"
    
    ' valid request?
    If vxmlRequestNode.selectSingleNode("BATCHAUDIT[@BATCHAUDITGUID][@BATCHNUMBER][@BATCHRUNNUMBER]") Is Nothing Then
        '... if not raise an error
        errThrowError strFunctionName, oeXMLMissingAttribute, "missing BATCHAUDIT details"
    End If
    
    Set xmlReqDoc = New FreeThreadedDOMDocument40
    xmlReqDoc.validateOnParse = False
    xmlReqDoc.setProperty "NewParser", True
    xmlReqDoc.async = False
    
    ' set up request
    Set xmlElem = xmlReqDoc.appendChild(vxmlRequestNode.cloneNode(False))
    xmlElem.setAttribute "OPERATION", "IncrementBatchScheduleSuccess"
    
    Set xmlElem = xmlReqDoc.createElement("BATCHSCHEDULE")
    xmlElem.setAttribute "BATCHNUMBER", vxmlRequestNode.selectSingleNode("BATCHAUDIT/@BATCHNUMBER").Text
    xmlElem.setAttribute "BATCHRUNNUMBER", vxmlRequestNode.selectSingleNode("BATCHAUDIT/@BATCHRUNNUMBER").Text
    xmlElem.setAttribute "BATCHAUDITGUID", vxmlRequestNode.selectSingleNode("BATCHAUDIT/@BATCHAUDITGUID").Text
    xmlReqDoc.firstChild.appendChild xmlElem
                    
    ' Call BatchScheduleBO.IncrementBatchScheduleSuccess
    Set omObj = GetObjectContext.CreateInstance(gstrBATCH_SCHEDULER_COMPONENT & ".BatchScheduleBO")
    strResponse = omObj.omBatchRequest(xmlReqDoc.xml)
    Set omObj = Nothing
    errCheckXMLResponse strResponse, True
    
    Set xmlRespDoc = New FreeThreadedDOMDocument40
    xmlRespDoc.validateOnParse = False
    xmlRespDoc.setProperty "NewParser", True
    xmlRespDoc.async = False
    xmlRespDoc.loadXML strResponse
    
    'ik_debug
    'xmlRespDoc.save "c:\omiga4Trace\IncrementBatchScheduleSuccessResponse.xml"
    
    If Not xmlRespDoc.selectSingleNode("RESPONSE/BATCHSCHEDULE[@LASTTRANSACTION='1']") Is Nothing Then
        ' complete batch
        xmlReqDoc.documentElement.setAttribute "OPERATION", "CompleteBatch"
    
        'ik_debug
        'xmlReqDoc.save "c:\omiga4Trace\CompleteBatchRequest.xml"
        
        Set omObj = GetObjectContext.CreateInstance(gstrPAYMENTPROCESSING & ".PaymentProcessingBO")
        strResponse = omObj.omPayProcRequest(xmlReqDoc.xml)
        Set omObj = Nothing
    
        'ik_debug
        'xmlRespDoc.loadXML strResponse
        'xmlRespDoc.save "c:\omiga4Trace\CompleteBatchResponse.xml"
        
        errCheckXMLResponse strResponse, True
    
    
    End If

ProcessCompletionsInterface_Exit:
    
    Set omObj = Nothing
    Set xmlElem = Nothing
    Set xmlReqDoc = Nothing
    Set xmlRespDoc = Nothing
    
    errCheckError strFunctionName, TypeName(Me)

End Sub

'Private Sub ProcessCompletionsInterface( _
'    ByVal vxmlRequestNode As IXMLDOMNode, _
'    ByVal vxmlResponseNode As IXMLDOMNode)
'' header ----------------------------------------------------------------------------------
'' description:
''   Generates completions request and passes it to external mortgage admin system
'' pass:
''   vxmlRequestNode  XML Input request specifying the data and operation
''   vxmlResponseNode XML Reponse to add the data to
'' return:       N/A
'' history:
''   MC  27/04/01 Created
''   RF  13/12/01 SYS3489 (Omiplus CC32) Call omToCBA (via OmiQueue)
''------------------------------------------------------------------------------------------
'On Error GoTo ProcessCompletionsInterface_Exit
'
'    Const strFunctionName As String = "ProcessCompletionsInterface"
'
'    Dim objRBDO As Object
'    'SYS4020 MDC 07/02/2002 - Make AdminRules late bound
'    Dim objAdminRules As Object 'AdminRulesBO
'    'SYS4020 MDC 07/02/2002 - End
'    'SYS2719 - OmiPlus 20 Part 2
'    Dim objMessageQueue As IomMessageQueue
'    'SYS2719 End
'
'    Dim xmlDoc As FreeThreadedDOMDocument40
'    Dim xmlNode As IXMLDOMNode
'    Dim xmlAppNode As IXMLDOMNode
'    ' PSC 20/09/2002 BMIDS00480
''    Dim xmlParentNode As IXMLDOMNode
'    Dim xmlTempRequest As IXMLDOMNode
'    Dim xmlTempResponse As IXMLDOMNode
'    ' PSC 20/09/2002 BMIDS00480
''    Dim xmlNodeList As IXMLDOMNodeList
'    Dim xmlPaymentNode As IXMLDOMNode
'    Dim xmlBatchAuditNode As IXMLDOMNode
'    'SYS2719 MDC 20/09/2001
'    ' PSC 20/09/2002 BMIDS00480 - Start
''    Dim xmlDisbPaymentNode As IXMLDOMNode
''    Dim xmlNumbersResponse As IXMLDOMNode
''    Dim xmlNumberNode As IXMLDOMNode
'    ' PSC 20/09/2002 BMIDS00480 - End
'    'SYS2719 End
'
'    Dim xmlMessage As IXMLDOMNode       ' PSC 20/09/2002 BMIDS00480
'    Dim xmlChild As IXMLDOMNode         ' PSC 20/09/2002 BMIDS00480
'
'    Dim strResponse As String
'    ' PSC 20/09/2002 BMIDS00480 - Start
''    Dim strValueId As String
''    Dim strROFValueId As String
''    Dim strCBValueId As String
''    Dim strValue As String
''    Dim lngValue As Long
'    ' PSC 20/09/2002 BMIDS00480 - End
'
'    Dim intValue As Integer
'    Dim strSeqNo As String
'    Dim strFileName As String
'    Dim strQueueName As String
'    Dim xmlAttrib As IXMLDOMAttribute
'
'    '------------------------------------------------------------------------------------------
'    ' Validate
'    '------------------------------------------------------------------------------------------
'
'    Set xmlPaymentNode = xmlGetMandatoryNode(vxmlRequestNode, "PAYMENTRECORD")
'    xmlCheckMandatoryAttribute xmlPaymentNode, "APPLICATIONNUMBER"
''    xmlCheckMandatoryAttribute xmlNode, "APPLICATIONFACTFINDNUMBER"
'    strSeqNo = xmlGetMandatoryAttributeText(xmlPaymentNode, "PAYMENTSEQUENCENUMBER")
'    Set xmlBatchAuditNode = xmlGetMandatoryNode(vxmlRequestNode, "BATCHAUDIT")
'
'    '------------------------------------------------------------------------------------------
'    ' Instantiate Objects
'    '------------------------------------------------------------------------------------------
'
'    Set objRBDO = m_objContext.CreateInstance(gstrREQUEST_BROKER_COMPONENT & ".OmRequestDO")
'    Set xmlDoc = New FreeThreadedDOMDocument40
'    xmlDoc.validateOnParse = False
'    xmlDoc.setProperty "NewParser", True
'    xmlDoc.async = False
'
'    '------------------------------------------------------------------------------------------
'    ' Setup the Request
'    '------------------------------------------------------------------------------------------
'
'    Set xmlTempRequest = vxmlRequestNode.cloneNode(False)
'    Set xmlAppNode = xmlDoc.createElement("APPLICATION")
'    xmlCopyMandatoryAttribute xmlPaymentNode, xmlAppNode, "APPLICATIONNUMBER"
'    'xmlCopyMandatoryAttribute xmlPaymentNode, xmlAppNode, "APPLICATIONFACTFINDNUMBER"
'    xmlSetAttributeValue xmlTempRequest, "COMBOLOOKUP", "Y"
'    'xmlSetAttributeValue xmlTempRequest, "OPERATION", ""
'
'    'SD 16/08/2005 Mar47 - start
'    Set xmlAttrib = xmlDoc.createAttribute("RB_TEMPLATE")
'    xmlAttrib.Text = "AdminRulesTemplate"
'    xmlTempRequest.Attributes.setNamedItem xmlAttrib
'    'SD 16/08/2005 Mar47 - end
'
'    ' PSC 20/09/2002 BMIDS00480 - Start
'    ' Get the completions schema name
'    Dim strCompSchemaName As String
'    strCompSchemaName = GetGlobalParamString("CompletionSchemaName")
'
'    ' If no completion schema found then use APPLICATION to get all data
'    If Len(Trim$(strCompSchemaName)) = 0 Then
'        strCompSchemaName = "COMPLETIONS"
'    End If
'
'    xmlSetAttributeValue xmlAppNode, "_SCHEMA_", strCompSchemaName
'    ' PSC 20/09/2002 BMIDS00480 - End
'
'    xmlTempRequest.appendChild xmlAppNode
'
'    '------------------------------------------------------------------------------------------
'    ' Call Request Broker
'    '------------------------------------------------------------------------------------------
'
'    strResponse = objRBDO.OmDataRequest(xmlTempRequest.xml)
'    errCheckXMLResponse strResponse, True, vxmlResponseNode
'    xmlDoc.loadXML strResponse
'    Set xmlTempResponse = xmlDoc.documentElement
'    'Check if a record has been found...
'    Set xmlNode = xmlTempResponse.selectSingleNode("RESPONSE/APPLICATION/_MESSAGE_")
'    If Not xmlNode Is Nothing Then
'        '... and if not raise an error
'        errThrowError strFunctionName, oeRecordNotFound
'    End If
'
'    '------------------------------------------------------------------------------------------
'    'Process Response
'    '------------------------------------------------------------------------------------------
'
'    ' PSC 20/09/2002 BMIDS00480 - Comment Out Code
'
''    'Check that this is correct payment type. If Return of Funds or Cancel Balance then exit without further action
''    Set xmlDisbPaymentNode = xmlGetMandatoryNode(xmlTempResponse, "APPLICATION/APPLICATIONFACTFIND/PAYMENTDETAILS/PAYMENTRECORD[@PAYMENTSEQUENCENUMBER='" & strSeqNo & "']/DISBURSEMENTPAYMENT")
''    strValue = xmlGetMandatoryAttributeText(xmlDisbPaymentNode, "PAYMENTTYPE")
''    'strValueId = GetFirstComboValueId("PaymentType", "I")
''    strROFValueId = GetFirstComboValueId("PaymentType", "N")
''    strCBValueId = GetFirstComboValueId("PaymentType", "NCB")
''    If strValue <> strROFValueId And strValue <> strCBValueId Then
''        'This is Initial Adv, Installment, Retention or incentive release so do interfacing
''
''        'Remove addresses that are not Home/Main or Correspondence
''        Set xmlNodeList = xmlTempResponse.selectNodes("APPLICATION/APPLICATIONFACTFIND/CUSTOMERROLE/CUSTOMER/CUSTOMERVERSION/CUSTOMERADDRESS[@ADDRESSTYPE != '1' and @ADDRESSTYPE != '2']")
''        For Each xmlNode In xmlNodeList
''            'Remove the ADDRESS node
''            Set xmlParentNode = xmlNode.parentNode
''            xmlParentNode.removeChild xmlNode
''            'If no addresses exist for this parent CUSTOMERADDRESS node, remove it also
''            If xmlParentNode.childNodes.length = 0 Then
''                xmlParentNode.parentNode.removeChild xmlParentNode
''            End If
''        Next
''
''        'Remove employments that are not current
''        Set xmlNodeList = xmlTempResponse.selectNodes("APPLICATION/APPLICATIONFACTFIND/CUSTOMERROLE/CUSTOMER/CUSTOMERVERSION/INCOME/EMPLOYMENT[@DATELEFTORCEASEDTRADING]")
''        For Each xmlNode In xmlNodeList
''            'Remove the EMPLOYMENT node
''            Set xmlParentNode = xmlNode.parentNode
''            xmlParentNode.removeChild xmlNode
''        Next
''
''        'Find APPLICATIONBANKBUILDINGSOC where PreferredPaymentDay not specified
''        Set xmlNodeList = xmlTempResponse.selectNodes("APPLICATION/APPLICATIONFACTFIND/APPLICATIONBANKBUILDINGSOC[not(@PREFERREDPAYMENTDAY)]")
''        If xmlNodeList.length > 0 Then
''            'Find default value for PreferredPaymentDay
''            lngValue = GetGlobalParamAmount("PreferredPaymentDay")
''            'Update the APPLICATIONBANKBUILDINGSOC accordingly
''            For Each xmlNode In xmlNodeList
''                xmlSetAttributeValue xmlNode, "PREFERREDPAYMENTDAY", CStr(lngValue)
''            Next
''        End If
'
'        'SYS2719 MDC 20/09/2001
'        '------------------------------------------------------------------------------------------
'        'Reserve Admin System Customer Numbers
'        '------------------------------------------------------------------------------------------
'
'        ' PSC 20/09/2002 BMIDS00480 - Comment Out Code
'
''        Set xmlParentNode = xmlDoc.createElement("NUMBERS")
''
''        'Customer
''        Set xmlNumberNode = GetCustomerNumberForCompletions("1")
''        Set xmlNode = xmlDoc.createElement("CUSTOMER")
''        xmlParentNode.appendChild xmlNode
''        xmlCopyMandatoryAttributeValue xmlNumberNode, xmlNode, "OTHERSYSTEMNUMBER", "NUMBER"
''
''        'Is this a joint application?
''        Set xmlNode = xmlGetMandatoryNode(xmlTempResponse, ".//APPLICATIONFACTFIND")
''        If xmlGetMandatoryAttributeAsLong(xmlNode, "NUMBEROFAPPLICANTS") > 1 Then
''            Set xmlNumberNode = GetCustomerNumberForCompletions("2")
''            Set xmlNode = xmlDoc.createElement("MORTGAGOR")
''            xmlParentNode.appendChild xmlNode
''            xmlCopyMandatoryAttributeValue xmlNumberNode, xmlNode, "OTHERSYSTEMNUMBER", "NUMBER"
''        End If
''
''        'Is the solicitor ad-hoc or panel?
''        Set xmlNode = xmlGetMandatoryNode(xmlTempResponse, ".//APPLICATIONFACTFIND/THIRDPARTYDATA/APPLICATIONLEGALREP")
''        If Not xmlAttributeValueExists(xmlNode, "DIRECTORYGUID") Then
''            Set xmlNumberNode = GetCustomerNumberForCompletions("3")
''            Set xmlNode = xmlDoc.createElement("SOLICITOR")
''            xmlParentNode.appendChild xmlNode
''            xmlCopyMandatoryAttributeValue xmlNumberNode, xmlNode, "OTHERSYSTEMNUMBER", "NUMBER"
''        End If
''
''        'Is the payee anything other than the Solicitor or Bank
''        intValue = xmlGetAttributeAsInteger(xmlDisbPaymentNode, "PAYEETYPE")
''        If Not (IsValidationType("PayeeType", intValue, "B") Or IsValidationType("PayeeType", intValue, "S")) Then
''            Set xmlNumberNode = GetCustomerNumberForCompletions("4")
''            Set xmlNode = xmlDoc.createElement("PAYEE")
''            xmlParentNode.appendChild xmlNode
''            xmlCopyMandatoryAttributeValue xmlNumberNode, xmlNode, "OTHERSYSTEMNUMBER", "NUMBER"
''        End If
'        'SYS2719 End
'
'        '------------------------------------------------------------------------------------------
'        'Call Rules component to transform XML to a valid OSG Request block
'        '------------------------------------------------------------------------------------------
'
'        Set xmlTempRequest = vxmlRequestNode.cloneNode(False)
'        xmlSetAttributeValue xmlTempRequest, "OPERATION", "RunCompletionRules"
'
'        ' PSC 20/09/2002 BMIDS00480 Comment Out Code
'        'SYS2719 MDC 20/09/2001.
''        xmlTempRequest.appendChild xmlParentNode    'Admin Customer Numbers
'        'SYS2719 End
'
'        xmlTempRequest.appendChild xmlPaymentNode
'        xmlTempRequest.appendChild xmlBatchAuditNode
'
'        'Mar47 SD 16/08/05, added global parameters and combos - start
'        For Each xmlNode In xmlDoc.selectSingleNode("RESPONSE").childNodes
'            xmlTempRequest.appendChild xmlNode.cloneNode(True)
'        Next
'
'        'Set xmlNode = xmlGetMandatoryNode(xmlTempResponse, "APPLICATION")
'        'xmlTempRequest.appendChild xmlNode
'        'Mar47 SD 16/08/05, added global parameters and combos - end
'
'        Set objAdminRules = m_objContext.CreateInstance("omAdminRules.AdminRulesBO")
'        strResponse = objAdminRules.OmRequest(xmlTempRequest.xml)
'
'
'
'        '------------------------------------------------------------------------------------------
'        ' Write CB Request to a file if required
'        '------------------------------------------------------------------------------------------
'
'        If xmlGetAttributeAsBoolean(vxmlRequestNode, "DEBUGOUTPUT", "0") Then
'            ' PSC 20/09/2002 BMIDS00480
'            strFileName = App.Path & "\CompletionsRequest" & Format$(Now, "yyyymmdd-hhmmss") & ".XML"
'            intValue = FreeFile
'            Open strFileName For Output As #intValue
'            Print #intValue, strResponse
'            Close #intValue
'        End If
'
'        errCheckXMLResponse strResponse, True
'
'        ' PSC 20/09/2002 BMIDS00480
'        Set xmlDoc = xmlLoad(strResponse, strFunctionName)
'
'
'        'SYS2719 Start - OmiPlus 20 Part 2
'
'        '------------------------------------------------------------------------------------------
'        ' Write Request to a Message Queue to be passed to CB Adapter
'        '------------------------------------------------------------------------------------------
'
'        Set objMessageQueue = m_objContext.CreateInstance(App.Title & ".omMessageQueue")
'
'        strQueueName = GetGlobalParamString("AdminSystemQueueName")
'
'        If Len(Trim$(strQueueName)) = 0 Then
'            errThrowError strFunctionName, oeInvalidParameter, _
'                "AdminSystemQueueName not found"
'        End If
'
'        Set xmlTempRequest = vxmlRequestNode.cloneNode(False)
'        xmlSetAttributeValue xmlTempRequest, "OPERATION", "SendToQueue"
'        Set xmlNode = xmlDoc.createElement("MESSAGEQUEUE")
'        xmlSetAttributeValue xmlNode, "QUEUENAME", strQueueName
'
'        ' RF SYS3489 For core CompletionsInterfaceProgId will refer to omToCBA
'        xmlSetAttributeValue xmlNode, "PROGID", _
'            GetGlobalParamString("CompletionsInterfaceProgId")
'
'        ' PSC 20/09/2002 BMIDS00480 - Start
'        Set xmlMessage = vxmlRequestNode.cloneNode(False)
'
'        For Each xmlChild In xmlDoc.documentElement.childNodes
'            xmlMessage.appendChild xmlChild
'        Next
'
'        xmlSetAttributeValue xmlNode, "XML", xmlMessage.xml
'        ' PSC 20/09/2002 BMIDS00480 - End
'
'        xmlTempRequest.appendChild xmlNode
'        Set xmlTempResponse = objMessageQueue.SendToQueue(xmlTempRequest)
'        'SYS2719 End
'
'    ' PSC 20/09/2002 BMIDS00480
''    End If
'
'ProcessCompletionsInterface_Exit:
'    Set objRBDO = Nothing
'    Set xmlDoc = Nothing
'    Set xmlNode = Nothing
'    ' PSC 20/09/2002 BMIDS00480
''    Set xmlParentNode = Nothing
'    Set xmlTempRequest = Nothing
'    Set xmlTempResponse = Nothing
'    ' PSC 20/09/2002 BMIDS00480
''    Set xmlNodeList = Nothing
'    Set objMessageQueue = Nothing
'    Set xmlMessage = Nothing        ' PSC 20/09/2002 BMIDS00480
'    Set xmlChild = Nothing          ' PSC 20/09/2002 BMIDS00480
'
'
'    errCheckError strFunctionName, TypeName(Me)
'
'End Sub


Private Function GetCustomerNumberForCompletions(ByVal strNumberRequired As String) As IXMLDOMNode

On Error GoTo GetCustomerNumberForCompletionsExit

Const strFunctionName As String = "GetCustomerNumberForCompletions"

Dim xmlDoc As FreeThreadedDOMDocument40
Dim xmlResponse As IXMLDOMNode
Dim xmlRequest As IXMLDOMNode
Dim xmlNode As IXMLDOMNode

    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    Set xmlResponse = xmlDoc.createElement("RESPONSE")
    Set xmlRequest = xmlDoc.createElement("REQUEST")
    
    Set xmlNode = xmlDoc.createElement("NUMBERREQUEST")
    xmlRequest.appendChild xmlNode
    xmlSetAttributeValue xmlNode, "NUMBERREQUIRED", strNumberRequired
    ' DM  18/02/2002 SYS4105 changed XML to match specification and the attribute within GetNewNumbers.
    xmlSetAttributeValue xmlNode, "NUMBERTYPEREQUIRED", "C"
    xmlSetAttributeValue xmlNode, "SOURCE", "Omiga"
    
    'DM 20/02/2002 SYS2721 Bodge to get newnumbers from the admin system to work
    'Call GetNewNumbers(xmlRequest, xmlResponse)
    Call GetNewNumbersWithAutoLogon(xmlRequest, xmlResponse)
    ' SYS2721 End
    
    Set GetCustomerNumberForCompletions = xmlGetMandatoryNode(xmlResponse, ".//OTHERSYSTEMNUMBER")

GetCustomerNumberForCompletionsExit:
    Set xmlDoc = Nothing
    Set xmlResponse = Nothing
    Set xmlRequest = Nothing
    Set xmlNode = Nothing

    errCheckError strFunctionName, TypeName(Me)

End Function

Private Function GetAccountSummary(ByVal vxmlRequestNode As IXMLDOMNode, _
                                   ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:  Gets the account summary for the account passed in
' pass:         vxmlRequestNode  XML Input request specifying the data and operation
'               vxmlResponseNode XML Reponse to add the data to
' return:       N/A
'------------------------------------------------------------------------------------------
On Error GoTo GetAccountSummaryExit

    Const cstrFunctionName As String = "GetAccountSummary"
        
    Dim xmlTransformerResponse As New FreeThreadedDOMDocument40
    xmlTransformerResponse.validateOnParse = False
    xmlTransformerResponse.setProperty "NewParser", True
    Dim xmlChildNode As IXMLDOMNode
    Dim objODITransformer As Object
    Dim strTransformerResponse As String
        
    Set objODITransformer = CreateObject("ODITransformer.ODITransformerBO")
    strTransformerResponse = objODITransformer.Request(vxmlRequestNode.xml)
    xmlTransformerResponse.loadXML (strTransformerResponse)
    errCheckXMLResponseNode xmlTransformerResponse.selectSingleNode("RESPONSE"), vxmlResponseNode, True
    
    For Each xmlChildNode In xmlTransformerResponse.selectSingleNode("RESPONSE").childNodes
        vxmlResponseNode.appendChild xmlChildNode.cloneNode(True)
    Next

GetAccountSummaryExit:
  
    errCheckError cstrFunctionName, TypeName(Me)

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

Private Function GetUserMortgageAdminDetails(ByVal vxmlRequestNode As IXMLDOMNode, _
                                            ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:  Gets user mortgage admin details
' pass:         vxmlRequestNode  XML Input request specifying the user specific detail
'               vxmlResponseNode XML Reponse to add the data to
' return:       N/A
'------------------------------------------------------------------------------------------
    
    Dim xmlChildNode As IXMLDOMNode
    Dim objODITransformer As Object
    Dim strTransformerResponse As String
    Dim xmlTransformerResponse As FreeThreadedDOMDocument40
    
    Const cstrFunctionName As String = "GetUserMortgageAdminDetails"
    
    On Error GoTo GetUserMortgageAdminDetailsExit
    
    'Create the ODITransformer BO.
    Set objODITransformer = CreateObject("ODITransformer.ODITransformerBO")
    
    'Delegate the request through to it.
    strTransformerResponse = objODITransformer.Request(vxmlRequestNode.xml)
    
    'Load and check the response for errors.
    Set xmlTransformerResponse = New FreeThreadedDOMDocument40
    xmlTransformerResponse.validateOnParse = False
    xmlTransformerResponse.setProperty "NewParser", True
    xmlTransformerResponse.loadXML (strTransformerResponse)
    errCheckXMLResponseNode xmlTransformerResponse.selectSingleNode("RESPONSE"), vxmlResponseNode, True

    'Aggregate multiple response elements.
    For Each xmlChildNode In xmlTransformerResponse.selectSingleNode("RESPONSE").childNodes
        vxmlResponseNode.appendChild xmlChildNode.cloneNode(True)
    Next
            
GetUserMortgageAdminDetailsExit:
    Set xmlChildNode = Nothing
    Set objODITransformer = Nothing
    Set xmlTransformerResponse = Nothing
   
    errCheckError cstrFunctionName, TypeName(Me)

End Function

Private Function SaveUserMortgageAdminDetails(ByVal vxmlRequestNode As IXMLDOMNode, _
                                            ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:  Delegate the call to the ODITransformer.
' pass:         vxmlRequestNode  XML Input request specifying the specific user detail
'               vxmlResponseNode XML Reponse to add the data to
' return:       N/A
'------------------------------------------------------------------------------------------
    
    Dim xmlChildNode As IXMLDOMNode
    Dim objODITransformer As Object
    Dim strTransformerResponse As String
    Dim xmlTransformerResponse As FreeThreadedDOMDocument40
    
    Const cstrFunctionName As String = "SaveUserMortgageAdminDetails"
    
    On Error GoTo GetUserMortgageAdminDetailsExit
    
    'Create the ODITransformer BO.
    Set objODITransformer = CreateObject("ODITransformer.ODITransformerBO")
    
    'Delegate the request through to it.
    strTransformerResponse = objODITransformer.Request(vxmlRequestNode.xml)
    
    'Load and check the response for errors.
    Set xmlTransformerResponse = New FreeThreadedDOMDocument40
    xmlTransformerResponse.validateOnParse = False
    xmlTransformerResponse.setProperty "NewParser", True
    xmlTransformerResponse.loadXML (strTransformerResponse)
    errCheckXMLResponseNode xmlTransformerResponse.selectSingleNode("RESPONSE"), vxmlResponseNode, True

    'Aggregate multiple response elements.
    For Each xmlChildNode In xmlTransformerResponse.selectSingleNode("RESPONSE").childNodes
        vxmlResponseNode.appendChild xmlChildNode.cloneNode(True)
    Next
            
GetUserMortgageAdminDetailsExit:
    Set xmlChildNode = Nothing
    Set objODITransformer = Nothing
    Set xmlTransformerResponse = Nothing
   
    errCheckError cstrFunctionName, TypeName(Me)
 
End Function

Private Sub SaveThirdPartyDetails(ByVal vxmlRequestNode As IXMLDOMNode, _
                                            ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:  Saves Thirdparty details
' pass:         vxmlRequestNode  XML Input request specifying the specific thirdparty details
'               vxmlResponseNode XML Reponse to add the data to
' return:       N/A
'------------------------------------------------------------------------------------------

On Error GoTo SaveThirdPartyDetailsExit

    Const cstrFunctionName As String = "SaveThirdPartyDetails"
    
    Dim xmlTransformerResponse As New FreeThreadedDOMDocument40
    xmlTransformerResponse.validateOnParse = False
    xmlTransformerResponse.setProperty "NewParser", True
    Dim xmlChildNode As IXMLDOMNode
    Dim objODITransformer As Object
    Dim strTransformerResponse As String
    
        
    Set objODITransformer = CreateObject("ODITransformer.ODITransformerBO")
    strTransformerResponse = objODITransformer.Request(vxmlRequestNode.xml)
    xmlTransformerResponse.loadXML (strTransformerResponse)
    errCheckXMLResponseNode xmlTransformerResponse.selectSingleNode("RESPONSE"), vxmlResponseNode, True
    
    For Each xmlChildNode In xmlTransformerResponse.selectSingleNode("RESPONSE").childNodes
        vxmlResponseNode.appendChild xmlChildNode.cloneNode(True)
    Next

SaveThirdPartyDetailsExit:
    Set xmlTransformerResponse = Nothing
    Set xmlChildNode = Nothing
    Set objODITransformer = Nothing
    
    errCheckError cstrFunctionName, TypeName(Me)

End Sub

Private Function SaveMortgageProductDetails(ByVal vxmlRequestNode As IXMLDOMNode, _
                                            ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:  Delegate the call to the ODITransformer.
' pass:         vxmlRequestNode  XML Input request specifying the specific user detail
'               vxmlResponseNode XML Reponse to add the data to
' return:       N/A
'------------------------------------------------------------------------------------------
    
    Dim xmlChildNode As IXMLDOMNode
    Dim objODITransformer As Object
    Dim strTransformerResponse As String
    Dim xmlTransformerResponse As FreeThreadedDOMDocument40
    
    Const cstrFunctionName As String = "SaveMortgageProductDetails"
    
    On Error GoTo SaveMortgageProductDetailsExit
    
    'Create the ODITransformer BO.
    Set objODITransformer = CreateObject("ODITransformer.ODITransformerBO")
    
    'Delegate the request through to it.
    strTransformerResponse = objODITransformer.Request(vxmlRequestNode.xml)
    
    'Load and check the response for errors.
    Set xmlTransformerResponse = New FreeThreadedDOMDocument40
    xmlTransformerResponse.validateOnParse = False
    xmlTransformerResponse.setProperty "NewParser", True
    xmlTransformerResponse.loadXML (strTransformerResponse)
    errCheckXMLResponseNode xmlTransformerResponse.selectSingleNode("RESPONSE"), vxmlResponseNode, True

    'Aggregate multiple response elements.
    For Each xmlChildNode In xmlTransformerResponse.selectSingleNode("RESPONSE").childNodes
        vxmlResponseNode.appendChild xmlChildNode.cloneNode(True)
    Next
            
SaveMortgageProductDetailsExit:
    Set xmlChildNode = Nothing
    Set objODITransformer = Nothing
    Set xmlTransformerResponse = Nothing
   
    errCheckError cstrFunctionName, TypeName(Me)
 
End Function

' PSC 31/10/2005 MAR352
Private Sub CallExternalSystem(ByVal vxmlRequestNode As IXMLDOMNode, _
                                      ByVal vxmlResponseNode As IXMLDOMNode, _
                                      ByVal vstrMessageType As String)
' header ----------------------------------------------------------------------------------
' description:
'   Calls the BMTranslationBridge to communicate with the BMids Admin system.
'   Optionally produces debug file in following format:
'   <Debug>
'       <TimePoints
'           RequestReady="hh:mm:ss"
'           BMTBCreated="hh:mm:ss"
'           ResponseReceived="hh:mm:ss" />
'       <TimeDifferences
'           OverallTime="hh:mm:ss"
'           ObjectCreation="hh:mm:ss"
'           BridgeProcessing="hh:mm:ss" />
'       <REQUEST> ... </REQUEST>
'       <RESPONSE> ... </RESPONSE>
'   </Debug>
' pass:
'   vxmlRequestNode
'       XML REQUEST node specifying the data/request to pass to the admin system
'   vxmlResponseNode
'       XML RESPONSE node containing the response from the admin system
' return: none
' history:
'                   BMIDS00026, BMIDS00056, BMIDS00025 Function to call BMids Admin System
'RF     10/06/2003  00008119    Add extra logging to CallBMidsAdminSystem - Add logging at point of creation of BMTranslationBridge.ExecuteBMService object
'RF     17/07/2003  00008267    Ensure CallBMidsAdminSystem logs requests even if no response is received
'------------------------------------------------------------------------------------------
On Error GoTo CallExternalSystemExit
    
    Const cstrFunctionName As String = "CallExternalSystem"
    
    Dim objExternalSystem       As Object           ' PSC 31/10/2005 MAR352
    Dim strResponse             As String
    Dim xmlResponse             As FreeThreadedDOMDocument40
    Dim xmlChildNode            As IXMLDOMNode
    Dim strSavePath             As String
    Dim strFileName             As String
    Dim blnIsDebuggingEnabled   As Boolean
    Dim domDebug                As FreeThreadedDOMDocument40
    Dim nodeDebugRoot           As IXMLDOMNode
    Dim nodeDebugTimePoints     As IXMLDOMNode
    Dim nodeDebugTimeDiffs      As IXMLDOMNode
    Dim dtReady                 As Date
    Dim dtObjCreated            As Date
    Dim dtRespReceived          As Date
    'EP2_87 GHun
    Dim strAdminSystemProgID            As String
    Dim strAdminSystemCallMethodName    As String
    'EP2_87 End
    
    ' PSC 31/10/2005 MAR352 - Start
    Dim strMessageType As String
    Dim lngErrorCode As Long
    Dim strDescription As String
    ' PSC 31/10/2005 MAR352 - End

    ' PSC 31/10/2005 MAR352 - Start
    strMessageType = xmlGetAttributeText(vxmlRequestNode, "MessageType", "")
    
    If Len(strMessageType) = 0 Then
        ' Set the MessageType so the admin system can identify what to do with the XML being sent to it
        xmlSetAttributeValue vxmlRequestNode, "MessageType", vstrMessageType
    End If
    ' PSC 31/10/2005 MAR352 - End
    
    'EP2_87 GHun
    strAdminSystemProgID = GetGlobalParamString("AdminSystemProgID")
    If Len(strAdminSystemProgID) = 0 Then
        errThrowError cstrFunctionName, "Missing GlobalParameter AdminSystemProgID "
    End If
    
    strAdminSystemCallMethodName = GetGlobalParamString("AdminSystemCallMethodName")
    If Len(strAdminSystemCallMethodName) = 0 Then
        errThrowError cstrFunctionName, "Missing GlobalParameter AdminSystemCallMethodName "
    End If
    'EP2_87 End
    
    'BMIDS00092 Save Request XML to file when debugging
    strSavePath = GetAdminDebugPath()
    If Len(strSavePath) > 0 Then
        blnIsDebuggingEnabled = True
        'RF 17/07/2003 00008267 Start
        ' PSC 31/10/2005 MAR352
         strFileName = strSavePath & vstrMessageType & _
            "-" & Format$(Now, "yyyymmdd_hhnnss") & ".xml"
        'RF 10/06/2003 00008119 Start
        ' Save Request XML to file
        'SaveNodeToFile vxmlRequestNode, strSavePath & vstrBMMessageType & "-REQUEST"
        Set domDebug = xmlCreateDOMObject
        Set nodeDebugRoot = domDebug.appendChild(domDebug.createElement("Debug"))
        Set nodeDebugTimePoints = nodeDebugRoot.appendChild(domDebug.createElement("TimePoints"))
        Set nodeDebugTimeDiffs = nodeDebugRoot.appendChild(domDebug.createElement("TimeDifferences"))
        nodeDebugRoot.appendChild vxmlRequestNode.cloneNode(True)
        dtReady = Now()
        xmlSetAttributeValue nodeDebugTimePoints, "RequestReady", Format$(dtReady, "hh:mm:ss")
        'RF 10/06/2003 00008119 End
        SaveNodeToFile domDebug.documentElement, strFileName
        'RF 17/07/2003 00008267 End
    Else
        blnIsDebuggingEnabled = False
    End If
    'BMIDS00092 End

    ' PSC 31/10/2005 MAR352
        
    'EP2_87 GHun
    'Set objExternalSystem = CreateObject("omAdminCS.AdminInterfaceBO")
    Set objExternalSystem = CreateObject(strAdminSystemProgID)
    'EP2_87 End
        
    'RF 10/06/2003 00008119 Start
    If blnIsDebuggingEnabled Then
        dtObjCreated = Now()
        ' PSC 31/10/2005 MAR352
        ' PSC 16/11/2005 MAR601
        xmlSetAttributeValue nodeDebugTimePoints, "AdminObjectCreated", Format$(dtObjCreated, "hh:mm:ss")
    End If
    'RF 10/06/2003 00008119 End
    
    'EP2_87 GHun
    'strResponse = objExternalSystem.ExecuteService(vxmlRequestNode.xml)
    strResponse = CallByName(objExternalSystem, strAdminSystemCallMethodName, VbMethod, vxmlRequestNode.xml)
    'EP2_87 End
    
    Set xmlResponse = xmlCreateDOMObject
    xmlResponse.loadXML strResponse
    
    'BMIDS00092 Save Response XML to file when debugging
    If blnIsDebuggingEnabled Then
        'RF 10/06/2003 00008119 Start
        dtRespReceived = Now()
        nodeDebugRoot.appendChild xmlResponse.documentElement.cloneNode(True)
        xmlSetAttributeValue nodeDebugTimePoints, "ResponseReceived", Format$(dtRespReceived, "hh:mm:ss")
        xmlSetAttributeValue nodeDebugTimeDiffs, "OverallTime", Format$(dtRespReceived - dtReady, "hh:mm:ss")
        xmlSetAttributeValue nodeDebugTimeDiffs, "ObjectCreation", Format$(dtObjCreated - dtReady, "hh:mm:ss")
        xmlSetAttributeValue nodeDebugTimeDiffs, "BridgeProcessing", Format$(dtRespReceived - dtObjCreated, "hh:mm:ss")
        'SaveNodeToFile xmlResponse.documentElement, strSavePath & vstrBMMessageType & "-RESPONSE"
        'RF 17/07/2003 00008267 Start
        'SaveNodeToFile domDebug.documentElement, strSavePath & vstrBMMessageType
        SaveNodeToFile domDebug.documentElement, strFileName
        'RF 10/06/2003 00008119 Start
        'RF 17/07/2003 00008267 End
    End If
    'BMIDS00092 End
    
    ' PSC 31/10/2005 MAR352 - Start
    lngErrorCode = errCheckXMLResponseNode(xmlResponse.documentElement, vxmlResponseNode, False)
    
    If lngErrorCode = 0 Then
        For Each xmlChildNode In xmlResponse.selectSingleNode("RESPONSE").childNodes
            vxmlResponseNode.appendChild xmlChildNode
        Next
    Else
        strDescription = xmlGetNodeText(xmlResponse, "RESPONSE/ERROR/DESCRIPTION")
        errThrowError cstrFunctionName, 4500, strDescription
    End If
    ' PSC 31/10/2005 MAR352 - End
    
CallExternalSystemExit:
    
    Set objExternalSystem = Nothing            ' PSC 31/10/2005 MAR352
    Set xmlResponse = Nothing
    Set xmlChildNode = Nothing
    'RF 10/06/2003 00008119 Start
    Set domDebug = Nothing
    'RF 10/06/2003 00008119 End
        
    errCheckError cstrFunctionName, TypeName(Me)
    
End Sub
'BMIDS00026, BMIDS00056, BMIDS00025 End

'BMIDS00092 GHun
Private Function GetAdminDebugPath() As String
' header ----------------------------------------------------------------------------------
' description:  Gets the path to save admin debug info to from the registry.
'               Creates the folder if it does not exist.
' pass:         none
' return:       A string containing a valid path name. Returns "" if the path does not exist.
'------------------------------------------------------------------------------------------
    
    Const strRegSection As String = "HKLM\SOFTWARE\Omiga4\System Configuration\"
    
    Dim objWshShell     As Object
    Dim strDebugPath    As String
       
    ' Ignore any errors that occur
    On Error Resume Next
    
    Set objWshShell = CreateObject("WScript.Shell")
    
    ' Read the path to save admin debugging info to from the registry
    strDebugPath = Trim(objWshShell.RegRead(strRegSection & "AdminDebugPath"))
    
    If Len(strDebugPath) > 0 Then
    
        If Right(strDebugPath, 1) <> "\" Then
            strDebugPath = strDebugPath & "\"
        End If
        
        If Len(Dir(strDebugPath, vbDirectory)) = 0 Then
            'Create the folder if it does not exist
            MkDir strDebugPath
            If Len(Dir(strDebugPath, vbDirectory)) = 0 Then
                ' The output folder does not exist and cannot be created
                strDebugPath = vbNullString
            End If
        End If
    End If
    
    GetAdminDebugPath = strDebugPath
    
    'Clear any errors that may have occurred as they can be ignored
    Err.Clear
End Function
'BMIDS00092 End

'BMIDS00092 GHun
'Private Sub SaveNodeToFile(ByVal vxmlNode As IXMLDOMNode, ByVal vstrFilePrefix As String)
Private Sub SaveNodeToFile(ByVal vxmlNode As IXMLDOMNode, ByVal vstrFileName As String)
' header ----------------------------------------------------------------------------------
' description:  Saves an XML node to a file
' pass:         vxmlNode        The XML node to save
'               vstrFileName    The name of the file to save to.
' return:       none
' history:
'RF     17/07/2003  00008267    Ensure CallBMidsAdminSystem logs requests even if no response is received
'------------------------------------------------------------------------------------------

    Dim xmlDoc      As FreeThreadedDOMDocument40
    'Dim strFileName As String
    
    'Ignore any errors that occur
    On Error Resume Next
    
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    xmlDoc.appendChild vxmlNode.cloneNode(True)
    
    'strFileName = vstrFilePrefix & "-" & Format(Now, "yyyymmdd_hhnnss") & ".xml"
    'xmlDoc.save strFileName
    xmlDoc.save vstrFileName
    
    Set xmlDoc = Nothing
End Sub
'BMIDS00092 End

' PSC 22/10/2002 BMIDS00676 - Start
Private Function RunGenericInterface(ByVal vxmlRequestNode As IXMLDOMNode, _
                                     ByVal vxmlResponseNode As IXMLDOMNode)
' header ----------------------------------------------------------------------------------
' description:  Calls the BMids generic interface
' pass:         vxmlRequestNode     XML REQUEST node specifying the data/request to pass to the admin system
'               vxmlResponseNode    XML RESPONSE node containing the response from the admin system
' return:       none
'------------------------------------------------------------------------------------------
    
    On Error GoTo RunGenericInterfaceExit
    
    Const cstrFunctionName As String = "RunGenericInterface"
    
    Dim xmlApplication As IXMLDOMNode
    
    Dim strTaskId As String
    
    Set xmlApplication = xmlGetMandatoryNode(vxmlRequestNode, "APPLICATION")
    strTaskId = xmlGetMandatoryAttributeText(xmlApplication, "TASKID")
    
    ' Remove unwanted attributes
    xmlApplication.Attributes.removeNamedItem "TASKID"
    vxmlRequestNode.Attributes.removeNamedItem "ADMINSYSTEMSTATE"
    vxmlRequestNode.Attributes.removeNamedItem "ACTION"
    vxmlRequestNode.Attributes.removeNamedItem "USERAUTHORITYLEVEL"
    
    CallExternalSystem vxmlRequestNode, vxmlResponseNode, strTaskId
    
    
RunGenericInterfaceExit:
    
    Set xmlApplication = Nothing
        
    errCheckError cstrFunctionName, TypeName(Me)
    
End Function
' PSC 22/10/2002 BMIDS00676 - End

