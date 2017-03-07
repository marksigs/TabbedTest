VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 1  'NoTransaction
END
Attribute VB_Name = "omWSInterfaceBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------
'Workfile:          $Workfile: $
'Current Version    $Revision: $
'Last Modified      $Modtime: $
'Modified By        $Author: $

'Prog   Date        Description
'IK     24/10/2005  created for Project MARS
'IK     26/10/2005  MAR314 - fix to FindMortgageProducts - drop extraneous RESPONSE node
'IK     26/10/2005  MAR314 - transform all Phase 1 component responses
'                   to remove 'mixed-node' TEXT & RAW attributes
'IK     02/11/2005  MAR375 - add ValidateXml
'GHun   25/10/2005  MAR301 Added CreateKFI
'IK     09/11/2005  MAR470 (ING 789) further modifications to AddNamespaceDirectives
'IK     14/11/2005  MAR418 add GetNextMessageSequenceNumber, RestoreFirstTitleInterface
'IK     28/11/2005  MAR714, modify AddNamespaceDirectives to test for xmlns attribute in parent node
'IK     01/12/2005  MAR771 transform GetDocumentList to restrict response
'GHun   31/01/2006  MAR1166 transform RunIncomeCalculation response
'RF     15/02/2006  MAR1258 Add timings to xml trace
'RF     15/02/2006  MAR1258 Ensure change linked to ClearQuest record
'IK     07/03/2006  EP198, EP199 simplify IngestionInterface
'IK     14/03/2006  EP2 - add namespace directives to response
'---------------------------------------------------------------------------------------------
'-- EPSOM changes
'
'SR     18/09/2006  1. Modified procedure ValidateXml. Add namespaces for all the nodes weven when
'                   validation is not required and return xml that includes namespaces. This needs
'                   to be improved so as not add namespace foe every node but the ones defined in
'                   msgCommontypes.xsd
'                   2. Modifed omRequest.  return XML with added namespaces.
'SR     18/09/2006  Do not use project name in the namespaces. Use 'Omiga.vertex.co.uk' in place of
'                   'Epsom.Omiga.vertex.co.uk'
'SR     19/09/2006  Modifed method GetNameSpaceURI. All namespaces contaon 'Omiga.vertex.co.uk' without
'                   any project name .
'IK     21/09/2006  add namespace directives to error responses
'IK     10/10/2006  (creeping) namespace directives via xslt
'SR     16/10/2006  New methods 'GetBrokerProfile' and 'BrokerRegistration'
'SR     18/10/2006  EP2_11 - Modified 'BrokerRegistration'
'IK     18/10/2006  EP2_11 - add FindIntroducers
'SR     19/10/2006  EP2_11 - add UpdateBrokerProfile
'SR     19/10/2006  EP2_11 - modified UpdateBrokerProfile
'PSC    07/11/2006  EP2_41 Added CalculateKFIFees
'IK     15/11/2006  EP2_113 - add namespace directives to REQUEST validation errors for FindIntroducers
'                           - trace REQUEST validation errors
'PSC    17/11/2006  EP2_41  Amend AddNodeNamespaceDirectives for CalculateKFIFees
'IK     21/11/2006  EP2_134 - namespace directives via .xslt for StopAndSaveFMA,SubmitFMA
'SR     04/12/2006  EP2_283 - Add default ChannelID in FindMortgageProducts
'SR     05/12/2006  EP2_242 - modified 'BrokerRegistration'
'SR     05/12/2006  EP2_282 - modified 'UpdateBrokerProfile'
'PE     12/12/2006  EP2_421 - Web specific combo items
'PSC    12/12/2006  EP2_434 - Add AssignCaseToPackager
'PE     13/12/2006  EP2_487 - Rollback GetCombo Web Service
'PE     14/12/2006  EP2_480 - Web specific combo items.
'PE     05/02/2007  EP2_1233 - Modify GetComboList node path.
'GHun   16/03/2007  EP2_2058 - Added CallHandleInterfaceResponse and changed omRequest to call it
'GHun   03/04/2007  EP2_2182 - Changed AssignCaseToPackager & CallHandleInterfaceResponse
'SR     10/04/2007  EP2_2257 - modified UpdateBrokerProfile. Update OmigaUser.ChangePasswordIndicator, if required.
'--------------------------------------------------------------------------------------------------
Option Explicit

Private gblnTraceTest As Boolean
Private gblnTraceFlag As Boolean
Private gstrTraceDateTime As String
Private gintTraceDateTimeQ As Integer
Private gstrTraceFolder As String
Private gblnValidateTest As Boolean
Private gblnValidateFlag As Boolean
Private gstrValidateXsdFolder As String

Public Function omRequest(ByVal vstrMethod As String, ByVal vstrRequest As String) As String
    Dim strRetVal As String
    Const cstrMethodName As String = "OmRequest"
    On Error GoTo omRequestVbErr
    
    ' RF 15/02/2006 MAR1258 Start - Add timings to xml trace
    Dim dtStart As Date
    Dim dtEnd As Date
    dtStart = Now()
    'RF 15/02/2006 MAR1258 End
    
    ValidateXml vstrMethod, vstrRequest, True
    
    Select Case vstrMethod
    
        Case "GetComboList"
            strRetVal = GetComboList(vstrRequest)
    
        Case "GetApplicationData"
            strRetVal = CrudBoInterface(vstrRequest, "transformGetApplicationDataRequest")
    
        Case "GetApplicationSummaryData"
            strRetVal = CrudBoInterface(vstrRequest, "transformGetApplicationSummaryDataRequest")
            
        Case "GetCaseTrackingData"
            strRetVal = GetCaseTrackingData(vstrRequest)
            
        Case "GetDocumentID"
            strRetVal = CrudBoInterface(vstrRequest, "transformGetDocumentIDRequest")
            
        Case "GetDocumentList"
            strRetVal = GetDocumentList(vstrRequest)
            
        Case "GetNextMessageSequenceNumber"
            strRetVal = GetNextMessageSequenceNumber(vstrRequest)
            
        Case "FindBusinessForCustomer"
            strRetVal = FindBusinessForCustomer(vstrRequest)
            
        Case "FindDirectoryList"
            strRetVal = FindDirectoryList(vstrRequest)
            
        Case "FindMortgageProducts"
            strRetVal = FindMortgageProducts(vstrRequest)
            
        Case "RestoreFirstTitleInterface"
            strRetVal = CrudBoInterface(vstrRequest, "transformRestoreFirstTitleInterfaceRequest")
            
        Case "RegisterDocumentToCase"
            strRetVal = CrudBoInterface(vstrRequest, "transformRegisterDocumentToCaseRequest")
            
        Case "RunIncomeCalculations"
            strRetVal = RunIncomeCalculations(vstrRequest)
            
        'ik EP2 14/03/2006
        Case "StopAndSaveAiP"
            strRetVal = IngestionInterface(vstrMethod, vstrRequest, "transformStopAndSaveAipRequest.xslt")
            
        Case "StopAndSaveFMA"
            strRetVal = IngestionInterface(vstrMethod, vstrRequest, "transformStopAndSaveFMARequest.xslt")
            
        Case "SubmitAiP"
            strRetVal = IngestionInterface(vstrMethod, vstrRequest, "transformSubmitAiPRequest.xslt")
            
        Case "SubmitFMA"
            strRetVal = IngestionInterface(vstrMethod, vstrRequest, "transformSubmitFMARequest.xslt")
        'ik EP2 14/03/2006 ends
    
        Case "ValidateQuote"
            strRetVal = CrudBoInterface(vstrRequest, "transformValidateQuoteRequest")
            
        'MAR301 GHun
        Case "CreateKFI"
            strRetVal = CreateKFI(vstrRequest)
        'MAR301 End
            
        Case "BrokerRegistration"  'SR EP2_11
            strRetVal = BrokerRegistration(vstrRequest)
        
        Case "GetBrokerProfile"  'SR EP2_11
            strRetVal = CrudBoInterface(vstrRequest, "transformGetBrokerProfileRequest")
             
        Case "UpdateBrokerProfile"  'SR EP2_11
           strRetVal = UpdateBrokerProfile(vstrRequest)
            
        ' ik EP2_11
        Case "FindIntroducers"
            strRetVal = CrudBoInterface(vstrRequest, "FindIntroducersRequest")
            strRetVal = TransformErrorResponse(strRetVal, "FindIntroducersResponse.xslt")
            
        ' PSC 07/11/2006 EP2_41 - Start
        Case "CalculateKFIFees"
            strRetVal = CalculateKFIFees(vstrRequest)
        ' PSC 07/11/2006 EP2_41 - End
        'PSC 12/12/2006 - EP2_434 Start
        Case "AssignCaseToPackager"
            strRetVal = AssignCaseToPackager(vstrRequest)   'EP2_2058 GHun
        'PSC 12/12/2006 - EP2_434 End
        Case Else
            Err.Raise oeCommandFailed, cstrMethodName, "no code for method: " & vstrMethod
    
    End Select
    
    If Err.Number = 0 Then
        ' EP2_113
        ' RF 15/02/2006 MAR1258 Start - Add timings to xml trace (nb this won't include validation time)
        ' dtEnd = Now()
        ' TraceMessages vstrRequest, strRetVal, dtStart, dtEnd
        'RF 15/02/2006 MAR1258 End
        ' EP2_113 ends
        ValidateXml vstrMethod, strRetVal, False
    End If
    
omRequestVbErr:
    
    If Err.Number <> 0 Then
        If Err.Source = "OMIGAERROR" Then
            strRetVal = Err.Description
        Else
            If Err.Source <> cstrMethodName Then
                If Err.Source = App.EXEName Then
                    Err.Source = cstrMethodName
                Else
                    Err.Source = cstrMethodName & "." & Err.Source
                End If
            End If
            Err.Source = App.EXEName & "." & TypeName(Me) & "." & Err.Source
            strRetVal = FormatError(vstrRequest)
            ' EP2_113
            Select Case vstrMethod
                Case "FindIntroducers"
                    strRetVal = TransformErrorResponse(strRetVal, "FindIntroducersResponse.xslt")
            End Select
            ' EP2_113_ends
        End If
    End If

'IK EP2_1 21/09/2006
omRequest = AddNameSpaceDirectives(vstrMethod, strRetVal)
Debug.Print "FINAL RESPONSE STRING = " & vbCrLf & strRetVal

' EP2_113
' RF 15/02/2006 MAR1258 Start - Add timings to xml trace (nb this won't include validation time)
dtEnd = Now()
'TraceMessages vstrRequest, omRequest
TraceMessages vstrRequest, strRetVal, dtStart, dtEnd
' EP2_113_ends

End Function

Private Function GetCaseTrackingData(ByVal vstrRequest As String) As String

    Const cstrMethodName As String = "GetCaseTrackingData"
    On Error GoTo GetCaseTrackingDataVbErr
    
    Dim xmlInDoc As DOMDocument40
    Set xmlInDoc = New DOMDocument40
    xmlInDoc.setProperty "NewParser", True
    xmlInDoc.async = False
    xmlInDoc.loadXML vstrRequest
    
    If xmlInDoc.parseError.errorCode <> 0 Then
        xmlParseError xmlInDoc.parseError
    End If
    
    xmlTransform xmlInDoc, "transformGetCaseTrackingDataRequest.xslt"
    
    Dim omigaObj As Object
    Set omigaObj = GetObjectContext.CreateInstance("omCaseTrack.omCaseTrackBO")
    GetCaseTrackingData = omigaObj.OmCaseTrackRequest(xmlInDoc.xml)
    
GetCaseTrackingDataVbErr:

    Set omigaObj = Nothing
    Set xmlInDoc = Nothing
    
    CheckError cstrMethodName

End Function
'SR EP2_11 - new method BrokerRegistration
Private Function BrokerRegistration(ByVal vstrRequest As String) As String

    Const cstrMethodName As String = "BrokerRegistration"
    On Error GoTo BrokerRegistrationVbErr
    
    Dim strRetVal As String, strIntroducerID As String
    Dim xmlNewRequest As IXMLDOMElement, xmlElem As IXMLDOMElement
    Dim xmlInDoc As DOMDocument40, xmlDoc As DOMDocument40
    Dim omCRUD As Object
    
    Set xmlInDoc = New DOMDocument40
    xmlInDoc.setProperty "NewParser", True
    xmlInDoc.async = False
    xmlInDoc.loadXML vstrRequest
    
    If xmlInDoc.parseError.errorCode <> 0 Then
        xmlParseError xmlInDoc.parseError
    End If
     
    strRetVal = CrudBoInterface(vstrRequest, "transformBrokerRegistrationRequest")
    
    Set xmlDoc = New DOMDocument40
    xmlDoc.setProperty "NewParser", True
    xmlDoc.async = False
    xmlDoc.loadXML strRetVal
    
    If xmlDoc.selectSingleNode("RESPONSE/@TYPE").Text = "SUCCESS" Then
        ' Call GetMethod and return this response
        strIntroducerID = xmlDoc.documentElement.Attributes.getNamedItem("INTRODUCERID").Text
        
        Set xmlNewRequest = xmlDoc.createElement("REQUEST")
        xmlNewRequest.setAttribute "USERID", xmlInDoc.documentElement.getAttribute("USERID")
        xmlNewRequest.setAttribute "UNITID", xmlInDoc.documentElement.getAttribute("UNITID")
        xmlNewRequest.setAttribute "CHANNELID", xmlInDoc.documentElement.getAttribute("CHANNELID")
        xmlNewRequest.setAttribute "SCHEMA_NAME", "WebServiceSchema"
        xmlNewRequest.setAttribute "ENTITY_REF", "BROKER_PROFILE"
        xmlNewRequest.setAttribute "CRUD_OP", "READ"
        xmlNewRequest.setAttribute "postProcRef", "transformGetBrokerProfileResponse.xslt"
        
        Set xmlElem = xmlInDoc.createElement("INTRODUCER")
        xmlElem.setAttribute "INTRODUCERID", strIntroducerID
        xmlNewRequest.appendChild xmlElem
    
        Set omCRUD = GetObjectContext.CreateInstance("omCRUD.omCRUDBO")
        strRetVal = omCRUD.omRequest(xmlNewRequest.xml)
    Else  'SR 05/12/2006 : EP2_242
         xmlTransform xmlDoc, "transformGetBrokerProfileResponse.xslt"
         strRetVal = xmlDoc.xml
    End If

    BrokerRegistration = strRetVal
    
BrokerRegistrationVbErr:

    Set xmlInDoc = Nothing
    Set xmlDoc = Nothing
    Set xmlNewRequest = Nothing
    Set xmlElem = Nothing
    Set omCRUD = Nothing
    
    CheckError cstrMethodName

End Function

'SR EP2_11 - new method UpdateBrokerProfile
Private Function UpdateBrokerProfile(ByVal vstrRequest As String) As String

   Const cstrMethodName As String = "BrokerRegistration"
   On Error GoTo UpdateBrokerProfileVbErr
   
    Dim strRetVal As String, strIntroducerID As String
    Dim xmlNewRequest As IXMLDOMElement, xmlElem As IXMLDOMElement
    Dim xmlInDoc As DOMDocument40, xmlDoc As DOMDocument40
    Dim omCRUD As Object
    
    Set xmlInDoc = New DOMDocument40
    xmlInDoc.setProperty "NewParser", True
    xmlInDoc.async = False
    xmlInDoc.loadXML vstrRequest

    If xmlInDoc.parseError.errorCode <> 0 Then
        xmlParseError xmlInDoc.parseError
    End If
    
    xmlTransform xmlInDoc, "transformUpdateBrokerProfileRequest.xslt"
    Set omCRUD = GetObjectContext.CreateInstance("omCRUD.omCRUDBO")
    strRetVal = omCRUD.omRequest(xmlInDoc.xml)
    
    Set xmlDoc = New DOMDocument40
    xmlDoc.setProperty "NewParser", True
    xmlDoc.async = False
    xmlDoc.loadXML strRetVal
    
    If xmlDoc.selectSingleNode("RESPONSE/@TYPE").Text = "SUCCESS" Then
        'SR 10/04/2007 : EP2_2257 - if new password was created, set OmigaUser.ChangePasswordIndicator = 0
        'Assumes that INTRODUCER node passed in has USERID attribute set - as done on web site
        Dim strUserId As String
        Dim blnUpdateIndicatorSuccess As Boolean
        Dim blnIndicatorUpdaterequired As Boolean
        If Not xmlInDoc.selectSingleNode("//INTRODUCER/@USERID") Is Nothing Then
            strUserId = xmlInDoc.selectSingleNode("//INTRODUCER/@USERID").Text
        Else
            strUserId = ""
        End If
        If Not xmlInDoc.selectSingleNode("//PASSWORD[@CRUD_OP='CREATE']") Is Nothing Then
            blnIndicatorUpdaterequired = True
        End If
        
        If blnIndicatorUpdaterequired And strUserId <> "" Then
            Set xmlNewRequest = xmlDoc.createElement("REQUEST")
            xmlNewRequest.setAttribute "USERID", xmlInDoc.documentElement.getAttribute("USERID")
            xmlNewRequest.setAttribute "UNITID", xmlInDoc.documentElement.getAttribute("UNITID")
            xmlNewRequest.setAttribute "CHANNELID", xmlInDoc.documentElement.getAttribute("CHANNELID")
            xmlNewRequest.setAttribute "CRUD_OP", "UPDATE"
            xmlNewRequest.setAttribute "ENTITY_REF", "OMIGAUSER"
            
            Set xmlElem = xmlInDoc.createElement("OMIGAUSER")
            xmlElem.setAttribute "USERID", strUserId
            xmlElem.setAttribute "CHANGEPASSWORDINDICATOR", "0"
            xmlNewRequest.appendChild xmlElem
            strRetVal = omCRUD.omRequest(xmlNewRequest.xml)
            xmlDoc.loadXML strRetVal
            If xmlDoc.selectSingleNode("RESPONSE/@TYPE").Text = "SUCCESS" Then
                blnUpdateIndicatorSuccess = True
            Else
                xmlTransform xmlDoc, "transformGetBrokerProfileResponse.xslt"
                strRetVal = xmlDoc.xml
            End If
        End If
        
        If (blnIndicatorUpdaterequired And strUserId <> "" And blnUpdateIndicatorSuccess) Or _
            blnIndicatorUpdaterequired = False Then
        'SR 10/04/2007 : EP2_2257 - End
            ' Call GetMethod and return this response
            strIntroducerID = xmlInDoc.documentElement.selectSingleNode("INTRODUCER").Attributes.getNamedItem("INTRODUCERID").Text
            
            Set xmlNewRequest = xmlDoc.createElement("REQUEST")
            xmlNewRequest.setAttribute "USERID", xmlInDoc.documentElement.getAttribute("USERID")
            xmlNewRequest.setAttribute "UNITID", xmlInDoc.documentElement.getAttribute("UNITID")
            xmlNewRequest.setAttribute "CHANNELID", xmlInDoc.documentElement.getAttribute("CHANNELID")
            xmlNewRequest.setAttribute "SCHEMA_NAME", "WebServiceSchema"
            xmlNewRequest.setAttribute "ENTITY_REF", "BROKER_PROFILE"
            xmlNewRequest.setAttribute "CRUD_OP", "READ"
            xmlNewRequest.setAttribute "postProcRef", "transformGetBrokerProfileResponse.xslt"
            
            Set xmlElem = xmlInDoc.createElement("INTRODUCER")
            xmlElem.setAttribute "INTRODUCERID", strIntroducerID
            xmlNewRequest.appendChild xmlElem
            strRetVal = omCRUD.omRequest(xmlNewRequest.xml)
        End If 'SR 10/04/2007 : EP2_2257
    Else 'SR 05/12/2006 : EP2_282
         xmlTransform xmlDoc, "transformGetBrokerProfileResponse.xslt"
         strRetVal = xmlDoc.xml
    End If
   
   UpdateBrokerProfile = strRetVal
   
UpdateBrokerProfileVbErr:
    Set xmlInDoc = Nothing
    Set xmlDoc = Nothing
    Set xmlNewRequest = Nothing
    Set xmlElem = Nothing
    Set omCRUD = Nothing
    
    CheckError cstrMethodName
End Function

Private Function GetDocumentList(ByVal vstrRequest As String) As String

    Const cstrMethodName As String = "GetDocumentList"
    On Error GoTo GetDocumentListVbErr
    
    Dim xmlInDoc As DOMDocument40
    Set xmlInDoc = New DOMDocument40
    xmlInDoc.setProperty "NewParser", True
    xmlInDoc.async = False
    xmlInDoc.loadXML vstrRequest
    
    If xmlInDoc.parseError.errorCode <> 0 Then
        xmlParseError xmlInDoc.parseError
    End If
    
    xmlTransform xmlInDoc, "transformGetDocumentListRequest.xslt"
    
    Dim omigaObj As Object
    Set omigaObj = GetObjectContext.CreateInstance("omPM.PrintManagerBO")
    'ik_MAR771_20051201
    xmlInDoc.loadXML omigaObj.omRequest(xmlInDoc.xml)
    If Not xmlInDoc.selectSingleNode("RESPONSE[@TYPE='SUCCESS']") Is Nothing Then
        xmlTransform xmlInDoc, "transformGetDocumentListResponse.xslt"
    End If
    GetDocumentList = xmlInDoc.xml
    'ik_MAR771_20051201_ends
    
GetDocumentListVbErr:

    Set omigaObj = Nothing
    Set xmlInDoc = Nothing
    
    CheckError cstrMethodName

End Function

Private Function GetNextMessageSequenceNumber(ByVal vstrRequest As String) As String

    Const cstrMethodName As String = "GetNextMessageSequenceNumber"
    On Error GoTo GetNextMessageSequenceNumberVbErr
    
    Dim xmlDoc As DOMDocument40
    Dim xmlElem As IXMLDOMElement
    Dim omigaObj As Object
    
    Dim strCrudResponse As String, _
    strNextNo As String
    
    Set xmlDoc = New DOMDocument40
    xmlDoc.setProperty "NewParser", True
    xmlDoc.async = False
    
    strCrudResponse = CrudBoInterface(vstrRequest, "transformGetNextMessageSequenceNumberRequest")
    
    xmlDoc.loadXML strCrudResponse
    
    If xmlDoc.selectSingleNode("RESPONSE[@TYPE='SUCCESS']") Is Nothing Then
        GetNextMessageSequenceNumber = strCrudResponse
    Else
    
        Set omigaObj = GetObjectContext.CreateInstance("omDG.DirectGatewayBO")
        strNextNo = omigaObj.GetNextMessageSequenceNumber("FirstTitleOutboundMSN")
        
        Set xmlElem = xmlDoc.createElement("MESSAGESEQUENCENUMBER")
        xmlElem.Text = strNextNo
        xmlDoc.documentElement.appendChild xmlElem
        
        GetNextMessageSequenceNumber = xmlDoc.xml
    
    End If
    
GetNextMessageSequenceNumberVbErr:

    Set omigaObj = Nothing
    Set xmlElem = Nothing
    Set xmlDoc = Nothing
    
    CheckError cstrMethodName

End Function

Private Function FindBusinessForCustomer(ByVal vstrRequest As String) As String

    Const cstrMethodName As String = "FindBusinessForCustomer"
    On Error GoTo FindBusinessForCustomerVbErr
    
    Dim xmlInDoc As DOMDocument40
    Set xmlInDoc = New DOMDocument40
    xmlInDoc.setProperty "NewParser", True
    xmlInDoc.async = False
    xmlInDoc.loadXML vstrRequest
    
    If xmlInDoc.parseError.errorCode <> 0 Then
        xmlParseError xmlInDoc.parseError
    End If
    
    xmlTransform xmlInDoc, "transformFindBusinessForCustomerRequest.xslt"
    
    Dim omigaObj As Object
    Set omigaObj = GetObjectContext.CreateInstance("omApp.ApplicationManagerBO")
    'ik_20051026_mar314
    xmlInDoc.loadXML omigaObj.FindBusinessForCustomer(xmlInDoc.xml)
    If Not xmlInDoc.selectSingleNode("RESPONSE[@TYPE='SUCCESS']") Is Nothing Then
        xmlTransform xmlInDoc, "transformFindBusinessForCustomerResponse.xslt"
    End If
    FindBusinessForCustomer = xmlInDoc.xml
    'ik_20051026_mar314_ends
    
FindBusinessForCustomerVbErr:

    Set omigaObj = Nothing
    Set xmlInDoc = Nothing
    
    CheckError cstrMethodName

End Function

Private Function FindDirectoryList(ByVal vstrRequest As String) As String

    Const cstrMethodName As String = "FindDirectoryList"
    On Error GoTo FindDirectoryListVbErr
    
    Dim xmlInDoc As DOMDocument40
    Set xmlInDoc = New DOMDocument40
    xmlInDoc.setProperty "NewParser", True
    xmlInDoc.async = False
    xmlInDoc.loadXML vstrRequest
    
    If xmlInDoc.parseError.errorCode <> 0 Then
        xmlParseError xmlInDoc.parseError
    End If
    
    xmlTransform xmlInDoc, "transformFindDirectoryListRequest.xslt"
    
    Dim omigaObj As Object
    Set omigaObj = GetObjectContext.CreateInstance("omTP.ThirdPartyBO")
    FindDirectoryList = omigaObj.FindDirectoryList(xmlInDoc.xml)
    
FindDirectoryListVbErr:

    Set omigaObj = Nothing
    Set xmlInDoc = Nothing
    
    CheckError cstrMethodName

End Function

Private Function FindMortgageProducts(ByVal vstrRequest As String) As String

    Const cstrMethodName As String = "FindMortgageProducts"
    On Error GoTo FindMortgageProductsVbErr
    
    Dim xmlInDoc As DOMDocument40
    Dim omCRUD As Object
    Dim xmlOutDoc As DOMDocument40
    Dim xmlElem As IXMLDOMElement
    
    Set xmlInDoc = New DOMDocument40
    xmlInDoc.setProperty "NewParser", True
    xmlInDoc.async = False
    xmlInDoc.loadXML vstrRequest
    
    If xmlInDoc.parseError.errorCode <> 0 Then
        xmlParseError xmlInDoc.parseError
    End If
    
    xmlTransform xmlInDoc, "transformFindMortgageProductsRequest.xslt"
    
    'SR 04/12/2006 : EP2_283 - Add channelId from global param, if not mentioned already
    If xmlInDoc.selectSingleNode("//DISTRIBUTIONCHANNELID") Is Nothing Then
        Set omCRUD = GetObjectContext.CreateInstance("omCRUD.omCRUDBO")
        Dim strReq As String, strResp As String
        strReq = "<REQUEST CRUD_OP=""READ"" ENTITY_REF=""GLOBALPARAMETER""><GLOBALPARAMETER NAME=""WEBCHANNELID""/></REQUEST>"
        strResp = omCRUD.omRequest(strReq)
        
        Set xmlOutDoc = New DOMDocument40
        xmlOutDoc.loadXML (strResp)
        Set xmlElem = xmlOutDoc.createElement("DISTRIBUTIONCHANNELID")
        xmlElem.Text = xmlOutDoc.selectSingleNode("//GLOBALPARAMETER/@AMOUNT").Text
        xmlInDoc.selectSingleNode("REQUEST/MORTGAGEPRODUCTREQUEST").appendChild xmlElem
    End If
    
    Dim omigaObj As Object
    Set omigaObj = GetObjectContext.CreateInstance("omAQ.ApplicationQuoteBO")
    
    'ik_20051026_mar314
    xmlInDoc.loadXML omigaObj.FindMortgageProducts(xmlInDoc.xml)
    If Not xmlInDoc.selectSingleNode("RESPONSE[@TYPE='SUCCESS']") Is Nothing Then
        xmlTransform xmlInDoc, "transformFindMortgageProductsResponse.xslt"
    End If
        
    FindMortgageProducts = xmlInDoc.xml
    'ik_20051026_mar314_ends
    
FindMortgageProductsVbErr:

    Set omigaObj = Nothing
    Set xmlInDoc = Nothing
    Set xmlOutDoc = Nothing
    Set omCRUD = Nothing
    Set xmlElem = Nothing
    
    CheckError cstrMethodName

End Function

Private Function CrudBoInterface( _
    ByVal vstrXmlIn As String, _
    ByVal vstrTransformerName As String) As String

    Const cstrMethodName As String = "CrudBoInterface"
    On Error GoTo CrudBoInterfaceVbErr
    
    Dim xmlInDoc As DOMDocument40
    Set xmlInDoc = New DOMDocument40
    xmlInDoc.setProperty "NewParser", True
    xmlInDoc.async = False
    xmlInDoc.loadXML vstrXmlIn
    
    If xmlInDoc.parseError.errorCode <> 0 Then
        xmlParseError xmlInDoc.parseError
    End If
    
    xmlTransform xmlInDoc, vstrTransformerName & ".xslt"
    
    Dim omCRUD As Object
    Set omCRUD = GetObjectContext.CreateInstance("omCRUD.omCRUDBO")
    CrudBoInterface = omCRUD.omRequest(xmlInDoc.xml)
    
CrudBoInterfaceVbErr:

    Set omCRUD = Nothing
    Set xmlInDoc = Nothing
    
    CheckError cstrMethodName

End Function

Private Function RunIncomeCalculations(ByVal vstrXmlIn As String) As String

    Const cstrMethodName As String = "RunIncomeCalculations"
    On Error GoTo RunIncomeCalculationsVbErr
    
    Dim xmlInDoc As DOMDocument40
    Set xmlInDoc = New DOMDocument40
    xmlInDoc.setProperty "NewParser", True
    xmlInDoc.async = False
    xmlInDoc.loadXML vstrXmlIn
    
    If xmlInDoc.parseError.errorCode <> 0 Then
        xmlParseError xmlInDoc.parseError
    End If
    
    xmlTransform xmlInDoc, "transformRunIncomeCalculationsRequest.xslt"
    
    Dim omigaObj As Object
    Set omigaObj = GetObjectContext.CreateInstance("omIC.IncomeCalcsBO")
    'MAR1166 GHun transform the response
    xmlInDoc.loadXML omigaObj.RunIncomeCalculation(xmlInDoc.xml)
    
    If Not xmlInDoc.selectSingleNode("RESPONSE[@TYPE='SUCCESS' or @TYPE='WARNING']") Is Nothing Then
        xmlTransform xmlInDoc, "transformRunIncomeCalculationsResponse.xslt"
    End If
    
    RunIncomeCalculations = xmlInDoc.xml
    'MAR1166 End
    
RunIncomeCalculationsVbErr:

    Set omigaObj = Nothing
    Set xmlInDoc = Nothing
    
    CheckError cstrMethodName

End Function
'ik EP2 14/03/2006
Private Function IngestionInterface( _
    ByVal vstrMethod As String, _
    ByVal vstrXmlIn As String, _
    ByVal vstrTransformerName As String) _
    As String

    Const cstrMethodName As String = "IngestionInterface"
    On Error GoTo IngestionInterfaceVbErr
    
    Dim xmlInDoc As DOMDocument40
    Dim objNode As IXMLDOMNode
    
    'ik EP2 14/03/2006
    Dim strFName As String
    
    Set xmlInDoc = New DOMDocument40
      
    xmlInDoc.setProperty "NewParser", True
    xmlInDoc.async = False
    xmlInDoc.loadXML vstrXmlIn
    
    If xmlInDoc.parseError.errorCode <> 0 Then
        xmlParseError xmlInDoc.parseError
    End If
    
    xmlTransform xmlInDoc, vstrTransformerName

    Dim omigaObj As Object
    Set omigaObj = GetObjectContext.CreateInstance("omIngestionManager.IngestionManagerBO")

    'ik EP2 14/03/2006
    Select Case xmlInDoc.documentElement.getAttribute("omigaClient")
    
        Case "epsom"
        
            xmlInDoc.loadXML omigaObj.omRequest(xmlInDoc.xml)
            
            strFName = "transform" & vstrMethod & "Response.xslt"
            xmlTransform xmlInDoc, strFName
            
            IngestionInterface = xmlInDoc.xml
        
        Case Else
    
            IngestionInterface = omigaObj.omRequest(xmlInDoc.xml)
            
    End Select
    
IngestionInterfaceVbErr:

    Set omigaObj = Nothing
    Set xmlInDoc = Nothing
    
    CheckError cstrMethodName

End Function

Private Sub ValidateXml( _
    ByVal vstrMethod As String, _
    ByVal vstrXml As String, _
    ByVal vblnIsRequest As Boolean)

    Const cstrMethodName As String = "ValidateXml"
    
    Dim xmlDoc As DOMDocument40
    Dim xsdDoc As XMLSchemaCache40
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode

    Dim strFileSpec As String, _
        strNsURI As String
    
    ' ik EP2 16/03/2006
    Dim blnIsValidateRequired As Boolean
    blnIsValidateRequired = IsValidateRequired(vstrXml, vblnIsRequest)
    
    If blnIsValidateRequired Then
    
        strFileSpec = gstrValidateXsdFolder & GetXsdFileSpec(vstrMethod, vblnIsRequest)
        strNsURI = GetNamespaceURI(vstrMethod, vblnIsRequest)
    
        Set xsdDoc = New XMLSchemaCache40
        xsdDoc.Add strNsURI, strFileSpec
        
        Set xmlDoc = New DOMDocument40
        xmlDoc.async = False
        xmlDoc.setProperty "NewParser", True
        ' ik EP2 16/03/2006
        xmlDoc.setProperty "SelectionNamespaces", "xmlns:ep='" & strNsURI & "'"
        Set xmlDoc.schemas = xsdDoc
        xmlDoc.validateOnParse = False
        
        xmlDoc.loadXML AddNameSpaceDirectives(vstrMethod, vstrXml)
        
        ' problems validating ERRORs skip for now
        ' ik EP2 16/03/2006 - epsom response does not have TYPE attribute
        If Not xmlDoc.selectSingleNode("ep:RESPONSE/ep:RESPONSESTATUS") Is Nothing Then
            If xmlDoc.selectSingleNode("ep:RESPONSE/ep:RESPONSESTATUS").Text <> "Success" Then
                blnIsValidateRequired = False
            End If
        End If
        
        If blnIsValidateRequired Then
        
            xmlDoc.validateOnParse = True
            xmlDoc.loadXML xmlDoc.xml
            
            If xmlDoc.parseError.errorCode <> 0 Then
                If vblnIsRequest Then
                    xmlValidationError xmlDoc.parseError, "XML validation error on REQUEST"
                Else
                    xmlValidationError xmlDoc.parseError, "XML validation error on RESPONSE"
                End If
            End If
            
        End If
    
    End If

ValidateXmlVbErr:
    CheckError cstrMethodName

End Sub

Private Function IsValidateRequired( _
    ByVal vstrXml As String, _
    ByVal vblnIsRequest As Boolean) _
    As Boolean
    
    Const cstrMethodName As String = "IsValidateRequired"
    
    Dim objWshShell As Object
    Dim fso As Scripting.FileSystemObject
    Dim xmlDoc As DOMDocument40
    Dim strRegValue As String

    On Error Resume Next
    
    If Not gblnValidateTest Then
    
        gblnValidateTest = True
    
        Set objWshShell = CreateObject("WScript.Shell")
        
        strRegValue = objWshShell.RegRead("HKLM\SOFTWARE\OMIGA4\wsValidate\")
        
        If strRegValue = "1" Then
            strRegValue = objWshShell.RegRead("HKLM\SOFTWARE\OMIGA4\wsValidate\xsdPath")
            If Len(strRegValue) = 0 Then
                App.LogEvent App.Title & "wsValidate - no xsd folder specified", vbLogEventTypeError
            Else
                Set fso = New Scripting.FileSystemObject
                If Not fso.FolderExists(strRegValue) Then
                    App.LogEvent App.Title & "wsValidate - cannot locate xsd folder: " & strRegValue, vbLogEventTypeError
                Else
                    If Right(strRegValue, 1) = "\" Then
                        gstrValidateXsdFolder = strRegValue
                    Else
                        gstrValidateXsdFolder = strRegValue & "\"
                    End If
                    gblnValidateFlag = True
                End If
            End If
        End If
        
        Set fso = Nothing
        Set objWshShell = Nothing
    
    End If
    
    On Error GoTo IsValidateRequiredVbErr
    
    ' ik EP2 16/03/2006
    
    IsValidateRequired = gblnValidateFlag

IsValidateRequiredVbErr:

    Set fso = Nothing
    Set objWshShell = Nothing
    Set xmlDoc = Nothing

    CheckError cstrMethodName

End Function

Private Function AddNameSpaceDirectives(ByVal vstrMethod As String, ByVal vstrXml As String) As String
    
     Dim xmlDoc As DOMDocument40
    Select Case vstrMethod
        
        Case "SubmitAiP"
            AddNameSpaceDirectives = vstrXml
        
        Case "GetBrokerProfile" 'SR EP2_11
            AddNameSpaceDirectives = vstrXml
        
        Case "BrokerRegistration" 'SR EP2_11
            AddNameSpaceDirectives = vstrXml
        
        Case "UpdateBrokerProfile" 'SR EP2_11
            AddNameSpaceDirectives = vstrXml
            
        ' ik EP2_11
        Case "FindIntroducers"
            AddNameSpaceDirectives = vstrXml
            
        ' ik EP2_134
        Case "StopAndSaveFMA", "SubmitFMA"
            AddNameSpaceDirectives = vstrXml
        
        Case Else
            Set xmlDoc = New DOMDocument40
            xmlDoc.async = False
            xmlDoc.setProperty "NewParser", True
            xmlDoc.validateOnParse = False
            
            xmlDoc.loadXML vstrXml
        
            AddNodeNamespaceDirectives vstrMethod, xmlDoc.documentElement
            AddNameSpaceDirectives = xmlDoc.xml
    
    End Select

    
    Set xmlDoc = Nothing

End Function

Private Sub AddNodeNamespaceDirectives( _
    ByVal vstrMethod As String, _
    ByVal vxmlNode As IXMLDOMElement)
    
    Const cstrMethodName As String = "AddNodeNamespaceDirectives"
    On Error GoTo AddNodeNamespaceDirectivesVbErr
    
    Dim xmlNode As IXMLDOMNode
    
    ' ik_MAR470_20051109
    If Not vxmlNode.Attributes.getNamedItem("xmlns") Is Nothing Then
        If Len(vxmlNode.Attributes.getNamedItem("xmlns").Text) = 0 Then
            vxmlNode.Attributes.removeNamedItem "xmlns"
        End If
    End If
    
    If vxmlNode.Attributes.getNamedItem("xmlns") Is Nothing Then
    ' ik_MAR470_20051109_ends
    
        Select Case vxmlNode.nodeName
            Case "REQUEST"
                ' will have namespace directive
            Case "RESPONSE"
                vxmlNode.setAttribute "xmlns", GetNamespaceURI(vstrMethod, False)
            Case "ADDRESS"
                Select Case vxmlNode.parentNode.baseName
                    Case "THIRDPARTY", "NAMEANDADDRESSDIRECTORY"
                        vxmlNode.setAttribute "xmlns", GetNamespaceURI("MsgTypes", False)
                    Case Else
                        Select Case vstrMethod
                            Case "StopAndSaveAiP", "SubmitAiP"
                                vxmlNode.setAttribute "xmlns", GetNamespaceURI("OmigaAiPData", False)
                            Case Else
                                vxmlNode.setAttribute "xmlns", GetNamespaceURI("OmigaData", False)
                        End Select
                End Select
            
            Case "CONTACTDETAILS", "CONTACTTELEPHONEDETAILS"
                vxmlNode.setAttribute "xmlns", GetNamespaceURI("MsgTypes", False)
            
            Case "APPLICATIONFACTFIND", "PAYMENTRECORD", "PAYEEHISTORY", "APPLICATIONFEETYPE"
                Select Case vstrMethod
                    'MAR301 GHun
                    ' PSC 17/11/2006 EP2_41
                    Case "CreateKFI", "CalculateKFIFees"
                        'ik_MAR714_20051128
                        If Not vxmlNode.parentNode.Attributes.getNamedItem("xmlns") Is Nothing Then
                            vxmlNode.setAttribute "xmlns", vxmlNode.parentNode.Attributes.getNamedItem("xmlns").Text
                        End If
                    'MAR301 End
                    Case "StopAndSaveAiP", "SubmitAiP"
                        vxmlNode.setAttribute "xmlns", GetNamespaceURI("OmigaAiPData", False)
                    Case Else
                        vxmlNode.setAttribute "xmlns", GetNamespaceURI("OmigaData", False)
                End Select
        End Select
    ' ik_MAR470_20051109
    End If
    
    ' IK EP2_1 22/09/2006
    If Not vxmlNode.parentNode.parentNode Is Nothing Then
        If vxmlNode.parentNode.parentNode.nodeName = "RESPONSE" Then
            If vxmlNode.parentNode.nodeName = "ERROR" _
            Or vxmlNode.parentNode.nodeName = "MESSAGE" _
            Then
                vxmlNode.setAttribute "xmlns", GetNamespaceURI("MsgTypes", False)
            End If
        End If
    End If
    
    If vxmlNode.Attributes.getNamedItem("xmlns") Is Nothing Then
        If Not vxmlNode.parentNode.Attributes.getNamedItem("xmlns") Is Nothing Then
            vxmlNode.setAttribute "xmlns", vxmlNode.parentNode.Attributes.getNamedItem("xmlns").Text
        End If
    End If
    
    For Each xmlNode In vxmlNode.childNodes
        If xmlNode.nodeType = NODE_ELEMENT Then
            AddNodeNamespaceDirectives vstrMethod, xmlNode
        End If
    Next
    
AddNodeNamespaceDirectivesVbErr:
    CheckError cstrMethodName
    
End Sub

Private Function GetXsdFileSpec( _
    ByVal vstrId As String, _
    ByVal vblnIsRequest As Boolean) _
    As String
    
    Const cstrCreateKFIRequest As String = "CreateKFIRequest.xsd"
    Const cstrCreateKFIResponse As String = "CreateKFIResponse.xsd"
    Const cstrFindBusinessForCustomerRequest As String = "FindBusinessForCustomerRequest.xsd"
    Const cstrFindBusinessForCustomerResponse As String = "FindBusinessForCustomerResponse.xsd"
    Const cstrFindDirectoryListRequest As String = "FindDirectoryListRequest.xsd"
    Const cstrFindDirectoryListResponse As String = "FindDirectoryListResponse.xsd"
    Const cstrFindMortgageProductsRequest As String = "FindMortgageProductsRequest.xsd"
    Const cstrFindMortgageProductsResponse As String = "FindMortgageProductsResponse.xsd"
    Const cstrFindStoredQuotesRequest As String = "FindStoredQuotesRequest.xsd"
    Const cstrFindStoredQuotesResponse As String = "FindStoredQuotesResponse.xsd"
    Const cstrGetApplicationDataRequest As String = "GetApplicationDataRequest.xsd"
    Const cstrGetApplicationDataResponse As String = "GetApplicationDataResponse.xsd"
    Const cstrGetApplicationSummaryDataRequest As String = "GetApplicationSummaryDataRequest.xsd"
    Const cstrGetApplicationSummaryDataResponse As String = "GetApplicationSummaryDataResponse.xsd"
    Const cstrGetApplicationSummaryForCustomerRequest As String = "GetApplicationSummaryForCustomerRequest.xsd"
    Const cstrGetApplicationSummaryForCustomerResponse As String = "GetApplicationSummaryForCustomerResponse.xsd"
    Const cstrGetCaseTrackingDataRequest As String = "GetCaseTrackingDataRequest.xsd"
    Const cstrGetCaseTrackingDataResponse As String = "GetCaseTrackingDataResponse.xsd"
    Const cstrGetComboListRequest As String = "GetComboListRequest.xsd"
    Const cstrGetComboListResponse As String = "GetComboListResponse.xsd"
    Const cstrgetCustomerDetailsRequest As String = "getCustomerDetailsRequest.xsd"
    Const cstrgetCustomerDetailsResponse As String = "getCustomerDetailsResponse.xsd"
    Const cstrGetDocumentIDRequest As String = "GetDocumentIDRequest.xsd"
    Const cstrGetDocumentIDResponse As String = "GetDocumentIDResponse.xsd"
    Const cstrGetDocumentListRequest As String = "GetDocumentListRequest.xsd"
    Const cstrGetDocumentListResponse As String = "GetDocumentListResponse.xsd"
    Const cstrPutCustomerDetailsRequest As String = "PutCustomerDetailsRequest.xsd"
    Const cstrPutCustomerDetailsResponse As String = "PutCustomerDetailsResponse.xsd"
    Const cstrRegisterDocumentToCaseRequest As String = "RegisterDocumentToCaseRequest.xsd"
    Const cstrRegisterDocumentToCaseResponse As String = "RegisterDocumentToCaseResponse.xsd"
    Const cstrRunIncomeCalculationsRequest As String = "RunIncomeCalculationsRequest.xsd"
    Const cstrRunIncomeCalculationsResponse As String = "RunIncomeCalculationsResponse.xsd"
    Const cstrSubmitAiPRequest As String = "SubmitAiPRequest.xsd"
    Const cstrSubmitAiPResponse As String = "SubmitAiPResponse.xsd"
    Const cstrSubmitFMARequest As String = "SubmitFMARequest.xsd"
    Const cstrSubmitFMAResponse As String = "SubmitFMAResponse.xsd"
    Const cstrSubmitStopAndSaveAiPRequest As String = "SubmitStopAndSaveAiPRequest.xsd"
    Const cstrSubmitStopAndSaveAiPResponse As String = "SubmitStopAndSaveAiPResponse.xsd"
    Const cstrSubmitStopAndSaveFMARequest As String = "SubmitStopAndSaveFMARequest.xsd"
    Const cstrSubmitStopAndSaveFMAResponse As String = "SubmitStopAndSaveFMAResponse.xsd"
    Const cstrValidateQuoteRequest As String = "ValidateQuoteRequest.xsd"
    Const cstrValidateQuoteResponse As String = "ValidateQuoteResponse.xsd"
    Const cstrRestoreFirstTitleInterfaceRequest As String = "RestoreFirstTitleInterfaceRequest.xsd"
    Const cstrRestoreFirstTitleInterfaceResponse As String = "RestoreFirstTitleInterfaceResponse.xsd"
    Const cstrGetNextMessageSequenceNumberRequest As String = "GetNextMessageSequenceNumberRequest.xsd"
    Const cstrGetNextMessageSequenceNumberResponse As String = "GetNextMessageSequenceNumberResponse.xsd"
    Const cstrBrokerRegistrationRequest As String = "BrokerRegAndMaintainRequest.xsd" 'SR EP2_11
    Const cstrBrokerRegistrationResponse As String = "BrokerRegAndMaintainResponse.xsd" 'SR EP2_11
    Const cstrGetBrokerProfileRequest As String = "BrokerRegAndMaintainRequest.xsd" 'SR EP2_11
    Const cstrGetBrokerProfileResponse As String = "BrokerRegAndMaintainResponse.xsd" 'SR EP2_11
    Const cstrUpdateBrokerProfileRequest As String = "BrokerRegAndMaintainRequest.xsd" 'SR EP2_11
    Const cstrUpdateBrokerProfileResponse As String = "BrokerRegAndMaintainResponse.xsd" 'SR EP2_11
    Const cstrFindIntroducersRequest As String = "FindIntroducersRequest.xsd" 'IK EP2_11
    Const cstrFindIntroducersResponse As String = "FindIntroducersResponse.xsd" 'IK EP2_11
    ' PSC 07/11/2006 EP2_41 - Start
    Const cstrCalculateKFIFeesRequest As String = "CalculateKFIFeesRequest.xsd"
    Const cstrCalculateKFIFeesResponse As String = "CalculateKFIFeesResponse.xsd"
    ' PSC 07/11/2006 EP2_41 - End
    'PSC 12/12/2006 EP2_434 - Start
    Const cstrAssignCaseToPackagerRequest As String = "AssignCaseToPackagerRequest.xsd"
    Const cstrAssignCaseToPackagerResponse As String = "AssignCaseToPackagerResponse.xsd"
    'PSC 12/12/2006 EP2_434 - End
    
    If vblnIsRequest Then
        Select Case vstrId
            'MAR301 GHun
            Case "CreateKFI"
                GetXsdFileSpec = cstrCreateKFIRequest
            'MAR301 End
            Case "GetComboList"
                GetXsdFileSpec = cstrGetComboListRequest
            Case "GetApplicationData"
                GetXsdFileSpec = cstrGetApplicationDataRequest
            Case "GetApplicationSummaryData"
                GetXsdFileSpec = cstrGetApplicationSummaryDataRequest
            Case "GetCaseTrackingData"
                GetXsdFileSpec = cstrGetCaseTrackingDataRequest
            Case "GetDocumentID"
                GetXsdFileSpec = cstrGetDocumentIDRequest
            Case "GetDocumentList"
                GetXsdFileSpec = cstrGetDocumentListRequest
            Case "FindBusinessForCustomer"
                GetXsdFileSpec = cstrFindBusinessForCustomerRequest
            Case "FindDirectoryList"
                GetXsdFileSpec = cstrFindDirectoryListRequest
            Case "FindMortgageProducts"
                GetXsdFileSpec = cstrFindMortgageProductsRequest
            Case "RegisterDocumentToCase"
                GetXsdFileSpec = cstrRegisterDocumentToCaseRequest
            Case "RunIncomeCalculations"
                GetXsdFileSpec = cstrRunIncomeCalculationsRequest
            Case "StopAndSaveAiP"
                GetXsdFileSpec = cstrSubmitStopAndSaveAiPRequest
            Case "StopAndSaveFMA"
                GetXsdFileSpec = cstrSubmitStopAndSaveFMARequest
            Case "SubmitAiP"
                GetXsdFileSpec = cstrSubmitAiPRequest
            Case "SubmitFMA"
                GetXsdFileSpec = cstrSubmitFMARequest
            Case "ValidateQuote"
                GetXsdFileSpec = cstrValidateQuoteRequest
            Case "RestoreFirstTitleInterface"
                GetXsdFileSpec = cstrRestoreFirstTitleInterfaceRequest
            Case "GetNextMessageSequenceNumber"
                GetXsdFileSpec = cstrGetNextMessageSequenceNumberRequest
            Case "BrokerRegistration"  'SR EP2_11
                GetXsdFileSpec = cstrBrokerRegistrationRequest
            Case "GetBrokerProfile"  'SR EP2_11
                GetXsdFileSpec = cstrGetBrokerProfileRequest
            Case "UpdateBrokerProfile"  'SR EP2_11
                GetXsdFileSpec = cstrUpdateBrokerProfileRequest
            Case "FindIntroducers"  'IK EP2_11
                GetXsdFileSpec = cstrFindIntroducersRequest
            ' PSC 07/11/2006 EP2_41 - Start
            Case "CalculateKFIFees"
                GetXsdFileSpec = cstrCalculateKFIFeesRequest
            ' PSC 07/11/2006 EP2_41 - End
            ' PSC 12/12/2006 EP2_434 - Start
            Case "AssignCaseToPackager"
                GetXsdFileSpec = cstrAssignCaseToPackagerRequest
            ' PSC 12/12/2006 EP2_434 - End
        End Select
    Else
        Select Case vstrId
            'MAR301 GHun
            Case "CreateKFI"
                GetXsdFileSpec = cstrCreateKFIResponse
            'MAR301 End
            Case "GetComboList"
                GetXsdFileSpec = cstrGetComboListResponse
            Case "GetApplicationData"
                GetXsdFileSpec = cstrGetApplicationDataResponse
            Case "GetApplicationSummaryData"
                GetXsdFileSpec = cstrGetApplicationSummaryDataResponse
            Case "GetCaseTrackingData"
                GetXsdFileSpec = cstrGetCaseTrackingDataResponse
            Case "GetDocumentID"
                GetXsdFileSpec = cstrGetDocumentIDResponse
            Case "GetDocumentList"
                GetXsdFileSpec = cstrGetDocumentListResponse
            Case "FindBusinessForCustomer"
                GetXsdFileSpec = cstrFindBusinessForCustomerResponse
            Case "FindDirectoryList"
                GetXsdFileSpec = cstrFindDirectoryListResponse
            Case "FindMortgageProducts"
                GetXsdFileSpec = cstrFindMortgageProductsResponse
            Case "RegisterDocumentToCase"
                GetXsdFileSpec = cstrRegisterDocumentToCaseResponse
            Case "RunIncomeCalculations"
                GetXsdFileSpec = cstrRunIncomeCalculationsResponse
            Case "StopAndSaveAiP"
                GetXsdFileSpec = cstrSubmitStopAndSaveAiPResponse
            Case "StopAndSaveFMA"
                GetXsdFileSpec = cstrSubmitStopAndSaveFMAResponse
            Case "SubmitAiP"
                GetXsdFileSpec = cstrSubmitAiPResponse
            Case "SubmitFMA"
                GetXsdFileSpec = cstrSubmitFMAResponse
            Case "ValidateQuote"
                GetXsdFileSpec = cstrValidateQuoteResponse
            Case "RestoreFirstTitleInterface"
                GetXsdFileSpec = cstrRestoreFirstTitleInterfaceResponse
            Case "GetNextMessageSequenceNumber"
                GetXsdFileSpec = cstrGetNextMessageSequenceNumberResponse
            Case "BrokerRegistration"  'SR EP2_11
                GetXsdFileSpec = cstrBrokerRegistrationResponse
            Case "GetBrokerProfile"  'SR EP2_11
                GetXsdFileSpec = cstrGetBrokerProfileResponse
            Case "UpdateBrokerProfile"  'SR EP2_11
                GetXsdFileSpec = cstrUpdateBrokerProfileResponse
            Case "FindIntroducers"  'IK EP2_11
                GetXsdFileSpec = cstrFindIntroducersResponse
            ' PSC 07/11/2006 EP2_41 - Start
            Case "CalculateKFIFees"
                GetXsdFileSpec = cstrCalculateKFIFeesResponse
            ' PSC 07/11/2006 EP2_41 - End
            ' PSC 12/12/2006 EP2_434 - Start
            Case "AssignCaseToPackager"
                GetXsdFileSpec = cstrAssignCaseToPackagerResponse
            ' PSC 12/12/2006 EP2_434 - End
        End Select
    End If
    
End Function


Private Function GetNamespaceURI( _
    ByVal vstrId As String, _
    ByVal vblnIsRequest As Boolean) _
    As String
    
    Const cstrCreateKFIRequest As String = "http://Request.CreateKFI.Omiga.vertex.co.uk"
    Const cstrCreateKFIResponse As String = "http://Response.CreateKFI.Omiga.vertex.co.uk"
    Const cstrFindBusinessForCustomerRequest As String = "http://Request.FindBusinessForCustomer.Omiga.vertex.co.uk"
    Const cstrFindBusinessForCustomerResponse As String = "http://Response.FindBusinessForCustomer.Omiga.vertex.co.uk"
    Const cstrFindDirectoryListRequest As String = "http://Request.FindDirectoryList.Omiga.vertex.co.uk"
    Const cstrFindDirectoryListResponse As String = "http://Response.FindDirectoryList.Omiga.vertex.co.uk"
    Const cstrFindMortgageProductsRequest As String = "http://Request.FindMortgageProducts.Omiga.vertex.co.uk"
    Const cstrFindMortgageProductsResponse As String = "http://Response.FindMortgageProducts.Omiga.vertex.co.uk"
    Const cstrFindStoredQuotesRequest As String = "http://Request.FindStoredQuotes.Omiga.vertex.co.uk"
    Const cstrFindStoredQuotesResponse As String = "http://Response.FindStoredQuotes.Omiga.vertex.co.uk"
    Const cstrGetApplicationDataRequest As String = "http://Request.GetApplicationData.Omiga.vertex.co.uk"
    Const cstrGetApplicationDataResponse As String = "http://Response.GetApplicationData.Omiga.vertex.co.uk"
    Const cstrGetApplicationSummaryDataRequest As String = "http://Request.GetApplicationSummaryData.Omiga.vertex.co.uk"
    Const cstrGetApplicationSummaryDataResponse As String = "http://Response.GetApplicationSummaryData.Omiga.vertex.co.uk"
    Const cstrGetApplicationSummaryForCustomerRequest As String = "http://Request.GetApplicationSummaryForCustomer.Omiga.vertex.co.uk"
    Const cstrGetApplicationSummaryForCustomerResponse As String = "http://Response.GetApplicationSummaryForCustomer.Omiga.vertex.co.uk"
    Const cstrGetCaseTrackingDataRequest As String = "http://Request.GetCaseTrackingData.Omiga.vertex.co.uk"
    Const cstrGetCaseTrackingDataResponse As String = "http://Response.GetCaseTrackingData.Omiga.vertex.co.uk"
    Const cstrGetComboListRequest As String = "http://Request.GetComboList.Omiga.vertex.co.uk"
    Const cstrGetComboListResponse As String = "http://Response.GetComboList.Omiga.vertex.co.uk"

    Const cstrGetDocumentIDRequest As String = "http://Request.GetDocumentID.Omiga.vertex.co.uk"
    Const cstrGetDocumentIDResponse As String = "http://Response.GetDocumentID.Omiga.vertex.co.uk"
    Const cstrGetDocumentListRequest As String = "http://Request.GetDocumentList.Omiga.vertex.co.uk"
    Const cstrGetDocumentListResponse As String = "http://Response.GetDocumentList.Omiga.vertex.co.uk"
    Const cstrRegisterDocumentToCaseRequest As String = "http://Request.RegisterDocumentToCase.Omiga.vertex.co.uk"
    Const cstrRegisterDocumentToCaseResponse As String = "http://Response.RegisterDocumentToCase.Omiga.vertex.co.uk"
    Const cstrRunIncomeCalculationsRequest As String = "http://Request.RunIncomeCalculations.Omiga.vertex.co.uk"
    Const cstrRunIncomeCalculationsResponse As String = "http://Response.RunIncomeCalculations.Omiga.vertex.co.uk"
    Const cstrSubmitAiPRequest As String = "http://Request.SubmitAiP.Omiga.vertex.co.uk"
    Const cstrSubmitAiPResponse As String = "http://Response.SubmitAiP.Omiga.vertex.co.uk"
    Const cstrSubmitFMARequest As String = "http://Request.SubmitFMA.Omiga.vertex.co.uk"
    Const cstrSubmitFMAResponse As String = "http://Response.SubmitFMA.Omiga.vertex.co.uk"
    Const cstrSubmitStopAndSaveAiPRequest As String = "http://Request.SubmitStopAndSaveAiP.Omiga.vertex.co.uk"
    Const cstrSubmitStopAndSaveAiPResponse As String = "http://Response.SubmitStopAndSaveAiP.Omiga.vertex.co.uk"
    Const cstrSubmitStopAndSaveFMARequest As String = "http://Request.SubmitStopAndSaveFMA.Omiga.vertex.co.uk"
    Const cstrSubmitStopAndSaveFMAResponse As String = "http://Response.SubmitStopAndSaveFMA.Omiga.vertex.co.uk"
    Const cstrValidateQuoteRequest As String = "http://Request.ValidateQuote.Omiga.vertex.co.uk"
    Const cstrValidateQuoteResponse As String = "http://Response.ValidateQuote.Omiga.vertex.co.uk"
    Const cstrRestoreFirstTitleInterfaceRequest As String = "http://Request.RestoreFirstTitleInterface.Omiga.vertex.co.uk"
    Const cstrRestoreFirstTitleInterfaceResponse As String = "http://Response.RestoreFirstTitleInterface.Omiga.vertex.co.uk"
    Const cstrGetNextMessageSequenceNumberRequest As String = "http://Request.GetNextMessageSequenceNumber.Omiga.vertex.co.uk"
    Const cstrGetNextMessageSequenceNumberResponse As String = "http://Response.GetNextMessageSequenceNumber.Omiga.vertex.co.uk"
    Const cstrBrokerRegistrationRequest As String = "http://Request.BrokerRegAndMaintain.Omiga.vertex.co.uk" 'SR EP2_11
    Const cstrBrokerRegistrationResponse As String = "http://Response.BrokerRegAndMaintain.Omiga.vertex.co.uk" 'SR EP2_11
    Const cstrGetBrokerProfileRequest As String = "http://Request.BrokerRegAndMaintain.Omiga.vertex.co.uk" 'SR EP2_11
    Const cstrGetBrokerProfileResponse As String = "http://Response.BrokerRegAndMaintain.Omiga.vertex.co.uk" 'SR EP2_11
    Const cstrFindIntroducersRequest As String = "http://Request.FindIntroducers.Omiga.vertex.co.uk" 'IK EP2_11
    Const cstrFindIntroducersResponse As String = "http://Response.FindIntroducers.Omiga.vertex.co.uk" 'IK EP2_11
    Const cstrUpdateBrokerProfileRequest As String = "http://Request.BrokerRegAndMaintain.Omiga.vertex.co.uk" 'SR EP2_11
    Const cstrUpdateBrokerProfileResponse As String = "http://Response.BrokerRegAndMaintain.Omiga.vertex.co.uk" 'SR EP2_11
    ' PSC 07/11/2006 EP2_41 - Start
    Const cstrCalculateKFIFeesRequest As String = "http://Request.CalculateKFIFees.Omiga.vertex.co.uk"
    Const cstrCalculateKFIFeesResponse As String = "http://Response.CalculateKFIFees.Omiga.vertex.co.uk"
    ' PSC 07/11/2006 EP2_41 - End
    ' PSC 12/12/2006 EP2_434 - Start
    Const cstrAssignCaseToPackagerRequest As String = "http://Request.AssignCaseToPackager.Omiga.vertex.co.uk"
    Const cstrAssignCaseToPackagerResponse As String = "http://Response.AssignCaseToPackager.Omiga.vertex.co.uk"
    ' PSC 12/12/2006 EP2_434 - End

    Const cstrOmigaAiPData As String = "http://OmigaAiPData.Omiga.vertex.co.uk"
    Const cstrOmigaData As String = "http://OmigaData.Omiga.vertex.co.uk"
    Const cstrMsgTypes As String = "http://msgtypes.Omiga.vertex.co.uk"
            
    If vblnIsRequest Then
        Select Case vstrId
            'MAR301 GHun
            Case "CreateKFI"
                GetNamespaceURI = cstrCreateKFIRequest
            'MAR301 End
            Case "GetComboList"
                GetNamespaceURI = cstrGetComboListRequest
            Case "GetApplicationData"
                GetNamespaceURI = cstrGetApplicationDataRequest
            Case "GetApplicationSummaryData"
                GetNamespaceURI = cstrGetApplicationSummaryDataRequest
            Case "GetCaseTrackingData"
                GetNamespaceURI = cstrGetCaseTrackingDataRequest
            Case "GetDocumentID"
                GetNamespaceURI = cstrGetDocumentIDRequest
            Case "GetDocumentList"
                GetNamespaceURI = cstrGetDocumentListRequest
            Case "FindBusinessForCustomer"
                GetNamespaceURI = cstrFindBusinessForCustomerRequest
            Case "FindDirectoryList"
                GetNamespaceURI = cstrFindDirectoryListRequest
            Case "FindMortgageProducts"
                GetNamespaceURI = cstrFindMortgageProductsRequest
            Case "RegisterDocumentToCase"
                GetNamespaceURI = cstrRegisterDocumentToCaseRequest
            Case "RunIncomeCalculations"
                GetNamespaceURI = cstrRunIncomeCalculationsRequest
            Case "StopAndSaveAiP"
                GetNamespaceURI = cstrSubmitStopAndSaveAiPRequest
            Case "StopAndSaveFMA"
                GetNamespaceURI = cstrSubmitStopAndSaveFMARequest
            Case "SubmitAiP"
                GetNamespaceURI = cstrSubmitAiPRequest
            Case "SubmitFMA"
                GetNamespaceURI = cstrSubmitFMARequest
            Case "ValidateQuote"
                GetNamespaceURI = cstrValidateQuoteRequest
            Case "RestoreFirstTitleInterface"
                GetNamespaceURI = cstrRestoreFirstTitleInterfaceRequest
            Case "GetNextMessageSequenceNumber"
                GetNamespaceURI = cstrGetNextMessageSequenceNumberRequest
            Case "BrokerRegistration"  'SR EP2_11
                GetNamespaceURI = cstrBrokerRegistrationRequest
            Case "GetBrokerProfile"  'SR EP2_11
                GetNamespaceURI = cstrGetBrokerProfileRequest
            Case "UpdateBrokerProfile"  'SR EP2_11
                GetNamespaceURI = cstrUpdateBrokerProfileRequest
            Case "FindIntroducers"  'IK EP2_11
                GetNamespaceURI = cstrFindIntroducersRequest
            ' PSC 07/11/2006 EP2_41 - Start
            Case "CalculateKFIFees"
                GetNamespaceURI = cstrCalculateKFIFeesRequest
            ' PSC 07/11/2006 EP2_41 - End
            ' PSC 12/12/2006 EP2_434 - Start
            Case "AssignCaseToPackager"
                GetNamespaceURI = cstrAssignCaseToPackagerRequest
            '  PSC 12/12/2006 EP2_434 - End
        End Select
    Else
        Select Case vstrId
            'MAR301 GHun
            Case "CreateKFI"
                GetNamespaceURI = cstrCreateKFIResponse
            'MAR301 End
            Case "GetComboList"
                GetNamespaceURI = cstrGetComboListResponse
            Case "GetApplicationData"
                GetNamespaceURI = cstrGetApplicationDataResponse
            Case "GetApplicationSummaryData"
                GetNamespaceURI = cstrGetApplicationSummaryDataResponse
            Case "GetCaseTrackingData"
                GetNamespaceURI = cstrGetCaseTrackingDataResponse
            Case "GetDocumentID"
                GetNamespaceURI = cstrGetDocumentIDResponse
            Case "GetDocumentList"
                GetNamespaceURI = cstrGetDocumentListResponse
            Case "FindBusinessForCustomer"
                GetNamespaceURI = cstrFindBusinessForCustomerResponse
            Case "FindDirectoryList"
                GetNamespaceURI = cstrFindDirectoryListResponse
            Case "FindMortgageProducts"
                GetNamespaceURI = cstrFindMortgageProductsResponse
            Case "RegisterDocumentToCase"
                GetNamespaceURI = cstrRegisterDocumentToCaseResponse
            Case "RunIncomeCalculations"
                GetNamespaceURI = cstrRunIncomeCalculationsResponse
            Case "StopAndSaveAiP"
                GetNamespaceURI = cstrSubmitStopAndSaveAiPResponse
            Case "StopAndSaveFMA"
                GetNamespaceURI = cstrSubmitStopAndSaveFMAResponse
            Case "SubmitAiP"
                GetNamespaceURI = cstrSubmitAiPResponse
            Case "SubmitFMA"
                GetNamespaceURI = cstrSubmitFMAResponse
            Case "ValidateQuote"
                GetNamespaceURI = cstrValidateQuoteResponse
            Case "RestoreFirstTitleInterface"
                GetNamespaceURI = cstrRestoreFirstTitleInterfaceResponse
            Case "GetNextMessageSequenceNumber"
                GetNamespaceURI = cstrGetNextMessageSequenceNumberResponse
            Case "BrokerRegistration"  'SR EP2_11
                GetNamespaceURI = cstrBrokerRegistrationResponse
            Case "GetBrokerProfile"  'SR EP2_11
                GetNamespaceURI = cstrGetBrokerProfileResponse
            Case "UpdateBrokerProfile"  'SR EP2_11
                GetNamespaceURI = cstrUpdateBrokerProfileResponse
            Case "FindIntroducers"  'IK EP2_11
                GetNamespaceURI = cstrFindIntroducersResponse
            Case "OmigaAiPData"
                GetNamespaceURI = cstrOmigaAiPData
            Case "OmigaData"
                GetNamespaceURI = cstrOmigaData
            Case "MsgTypes"
                GetNamespaceURI = cstrMsgTypes
            ' PSC 07/11/2006 EP2_41 - Start
            Case "CalculateKFIFees"
                GetNamespaceURI = cstrCalculateKFIFeesResponse
            ' PSC 07/11/2006 EP2_41 - End
            ' PSC 12/12/2006 EP2_434 - Start
            Case "AssignCaseToPackager"
                GetNamespaceURI = cstrAssignCaseToPackagerResponse
            '  PSC 12/12/2006 EP2_434 - End
        End Select
    End If

End Function

' RF 15/02/2006 MAR1258 Add timings to xml trace
Private Sub TraceMessages( _
    ByVal vstrRequest As String, _
    ByVal vstrResponse As String, _
    ByVal vdtStart As Date, _
    ByVal vdtEnd As Date)
    
    Dim fso As Scripting.FileSystemObject
    Dim objWshShell As Object
    
    Dim xmlTempDoc As DOMDocument40
    Dim xmlTraceDoc As DOMDocument40
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode

    Dim strRegValue As String, _
        strTempName As String, _
        strRef As String, _
        strTraceDateTime As String
        
    On Error Resume Next
    
    If Not gblnTraceTest Then
    
        gblnTraceTest = True
    
        Set objWshShell = CreateObject("WScript.Shell")
        
        strRegValue = objWshShell.RegRead("HKLM\SOFTWARE\OMIGA4\TRACE\om4Wrapper")
        
        If strRegValue = "1" Then
            
            strRegValue = objWshShell.RegRead("HKLM\SOFTWARE\OMIGA4\TRACE\FOLDER")
            If Len(strRegValue) = 0 Then
                App.LogEvent App.Title & "tracing - no trace folder specified", vbLogEventTypeError
            Else
                Set fso = New Scripting.FileSystemObject
                
                If Not fso.FolderExists(strRegValue) Then
                    fso.CreateFolder strRegValue
                End If
                
                If Not fso.FolderExists(strRegValue) Then
                    App.LogEvent App.Title & "tracing - cannot create trace folder: " & strRegValue, vbLogEventTypeError
                Else
                    If Right(strRegValue, 1) = "\" Then
                        gstrTraceFolder = strRegValue
                    Else
                        gstrTraceFolder = strRegValue & "\"
                    End If
                    gblnTraceFlag = True
                End If
                
                Set fso = Nothing
                
            End If
        
        End If
        
        Set objWshShell = Nothing
    
    End If
    
    If gblnTraceFlag Then
    
        Set xmlTraceDoc = New DOMDocument40
        xmlTraceDoc.setProperty "NewParser", True
        xmlTraceDoc.async = False
        
        Set xmlElem = xmlTraceDoc.createElement("om4MESSAGEPAIR")
        ' RF 15/02/2006 MAR1258 Start - Add timings to xml trace
        xmlElem.setAttribute "StartTime", Format(vdtStart, "dd/mm/yyyy hh:mm:ss")
        xmlElem.setAttribute "EndTime", Format(vdtEnd, "dd/mm/yyyy hh:mm:ss")
        ' RF 15/02/2006 MAR1258 End
        Set xmlNode = xmlTraceDoc.appendChild(xmlElem)
    
        Set xmlTempDoc = New DOMDocument40
        xmlTempDoc.setProperty "NewParser", True
        xmlTempDoc.async = False
        
        xmlTempDoc.loadXML vstrRequest
        
        If xmlTempDoc.parseError.errorCode = 0 Then
            xmlNode.appendChild xmlTempDoc.documentElement.cloneNode(True)
        Else
            Set xmlElem = xmlTraceDoc.createElement("REQUEST")
            xmlElem.Text = vstrRequest
            xmlNode.appendChild xmlElem
        End If
            
        xmlTempDoc.loadXML vstrResponse
        
        If xmlTempDoc.parseError.errorCode = 0 Then
            xmlNode.appendChild xmlTempDoc.documentElement.cloneNode(True)
        Else
            Set xmlElem = xmlTraceDoc.createElement("RESPONSE")
            xmlElem.Text = vstrResponse
            xmlNode.appendChild xmlElem
        End If
    
        If Not xmlTraceDoc.documentElement.firstChild.Attributes.getNamedItem("USERID") Is Nothing Then
            strRef = "_" & xmlTraceDoc.documentElement.firstChild.Attributes.getNamedItem("USERID").Text & "_"
        Else
            strRef = "_"
        End If
        
        strTraceDateTime = Format(Now(), "yyyymmdd_hhmmss")
        
        If strTraceDateTime = gstrTraceDateTime Then
            gintTraceDateTimeQ = gintTraceDateTimeQ + 1
        Else
            gstrTraceDateTime = strTraceDateTime
            gintTraceDateTimeQ = 1
        End If
    
        strTempName = gstrTraceFolder & "\" & "om4WrapperTrace" & strRef & strTraceDateTime & "_" & Right(CStr(gintTraceDateTimeQ + 1000), 3) & ".xml"
        
        xmlTraceDoc.save strTempName
        
        Set xmlNode = Nothing
        Set xmlElem = Nothing
        Set xmlTempDoc = Nothing
        Set xmlTraceDoc = Nothing

    End If
    
    On Error GoTo 0

End Sub

Private Sub xmlParseError(ByVal objParseError As IXMLDOMParseError)
    Err.Raise oeXMLParserError, "", "XML parser error - " & vbCr & FormatXmlError(objParseError)
End Sub

Private Sub xmlValidationError(ByVal objParseError As IXMLDOMParseError, ByVal vstrPrefix As String)
    Err.Raise oeXMLParserError, "", vstrPrefix & vbCr & FormatXmlError(objParseError)
End Sub

Private Function FormatXmlError(ByVal objParseError As IXMLDOMParseError) As String
    
    Dim strErrDesc As String    ' formatted parser error
    
    strErrDesc = _
        "Reason: " & objParseError.reason & vbCr & _
        "Error code: " & Str$(objParseError.errorCode) & vbCr & _
        "Line no.: " & Str$(objParseError.Line) & vbCr & _
        "Character: " & Str$(objParseError.linepos) & vbCr
        
    If Len(objParseError.srcText) <> 0 Then
        strErrDesc = vbCr & "Source text: " & objParseError.srcText
    End If

    FormatXmlError = strErrDesc

End Function

Private Sub xmlTransform(ByVal xmlDoc As DOMDocument40, ByVal vstrFileName)

    Dim xslDoc As DOMDocument40
    
    Dim strFileSpec As String
    
    strFileSpec = App.Path & "\" & vstrFileName
    strFileSpec = Replace(strFileSpec, "DLL", "XML", 1, 1, vbTextCompare)
    
    Set xslDoc = New DOMDocument40
    xslDoc.async = False
    xslDoc.Load strFileSpec
    
    If xslDoc.parseError.errorCode <> 0 Then
        Err.Raise _
            oeXMLParserError, "", _
                "error loading xslt document: " & _
                strFileSpec & _
                " " & xslDoc.parseError.reason
    End If
    
    xmlDoc.loadXML xmlDoc.transformNode(xslDoc)
    
    If xslDoc.parseError.errorCode <> 0 Then
        Err.Raise _
            oeXMLParserError, "", _
                "error in transformation: " & _
                strFileSpec & _
                " " & xslDoc.parseError.reason
    End If

End Sub

Private Sub CheckError(ByVal vstrMethodName As String)
    If Err.Number = 0 Then
        Exit Sub
    End If
    If Err.Source = "OMIGAERROR" Then
        ChainErrorSource vstrMethodName
    Else
        If Err.Source <> vstrMethodName Then
            If InStr(Err.Source, "." & vstrMethodName) = 0 Then
                If Err.Source = App.EXEName Then
                    Err.Source = vstrMethodName
                Else
                    Err.Source = vstrMethodName & "." & Err.Source
                End If
            End If
        End If
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Function ChainErrorSource(ByVal vstrMethodName As String) As String
    
    Dim xmlErrDoc As DOMDocument40
    Dim strSource As String, _
        strErrDesc As String, _
        strErrSource As String
    Dim lngErrNo As Long
    
    lngErrNo = Err.Number
    strErrDesc = Err.Description
    strErrSource = Err.Source
    
    On Error GoTo 0
    
    Set xmlErrDoc = New DOMDocument40
    xmlErrDoc.setProperty "NewParser", True
    xmlErrDoc.async = False
    xmlErrDoc.loadXML strErrDesc
    
    If Not xmlErrDoc.selectSingleNode("RESPONSE/ERROR/SOURCE") Is Nothing Then
        strSource = vstrMethodName & "." & xmlErrDoc.selectSingleNode("RESPONSE/ERROR/SOURCE").Text
        xmlErrDoc.selectSingleNode("RESPONSE/ERROR/SOURCE").Text = strSource
    End If
    
    Err.Number = lngErrNo
    Err.Source = strErrSource
    Err.Description = xmlErrDoc.xml
        
    Set xmlErrDoc = Nothing
    
End Function

Private Function FormatError(ByVal vstrXmlIn As String) As String
    
    Dim xmlErrDoc As DOMDocument40
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    
    Set xmlErrDoc = New DOMDocument40
    xmlErrDoc.setProperty "NewParser", True
    xmlErrDoc.async = False
    
    Set xmlElem = xmlErrDoc.createElement("RESPONSE")
    
    If Err.Source = "om4Wrapper" Then
        xmlElem.setAttribute "TYPE", "APPERR"
    Else
        xmlElem.setAttribute "TYPE", "SYSERR"
    End If
    
    Set xmlNode = xmlErrDoc.appendChild(xmlElem)
    
    Set xmlElem = xmlErrDoc.createElement("ERROR")
    Set xmlNode = xmlNode.appendChild(xmlElem)
    
    Set xmlElem = xmlErrDoc.createElement("NUMBER")
    xmlElem.Text = Err.Number
    xmlNode.appendChild xmlElem
    
    Set xmlElem = xmlErrDoc.createElement("SOURCE")
    xmlElem.Text = Err.Source
    xmlNode.appendChild xmlElem
    
    Set xmlElem = xmlErrDoc.createElement("VERSION")
    If Len(App.Comments) > 0 Then
        xmlElem.Text = App.Comments
    Else
        xmlElem.Text = App.Major & "." & App.Major & "." & App.Revision
    End If
    xmlNode.appendChild xmlElem
    
    Set xmlElem = xmlErrDoc.createElement("DESCRIPTION")
    xmlElem.Text = Err.Description
    xmlNode.appendChild xmlElem
    
    Set xmlElem = xmlErrDoc.createElement("REQUEST")
    xmlElem.Text = vstrXmlIn
    xmlNode.appendChild xmlElem
    
    FormatError = xmlErrDoc.xml
    
    Set xmlNode = Nothing
    Set xmlElem = Nothing
    Set xmlErrDoc = Nothing

End Function

'MAR301 GHun
Private Function CreateKFI(ByVal vstrRequest As String) As String
    Const cstrMethodName As String = "CreateKFI"
    Dim omigaObj As Object
    Dim xmlInDoc As DOMDocument40
    Dim objContext As ObjectContext
    
    On Error GoTo CreateKFIVbErr
    
    Set xmlInDoc = New DOMDocument40
    xmlInDoc.setProperty "NewParser", True
    xmlInDoc.async = False
    xmlInDoc.loadXML vstrRequest
    
    If xmlInDoc.parseError.errorCode <> 0 Then
        xmlParseError xmlInDoc.parseError
    End If
    
    xmlTransform xmlInDoc, "transformCreateKFIRequest.xslt"
    
    Set objContext = GetObjectContext()
    If objContext Is Nothing Then
        Set omigaObj = CreateObject("omCK.CreateKFIBO")
    Else
        Set omigaObj = objContext.CreateInstance("omCK.CreateKFIBO")
    End If

    CreateKFI = omigaObj.CreateKFI(xmlInDoc.xml)
    
CreateKFIVbErr:

    Set omigaObj = Nothing
    Set xmlInDoc = Nothing
    
    CheckError cstrMethodName
End Function
'MAR301 End

' ik EP2_11
Private Function TransformErrorResponse( _
    ByVal vxmlResponse As String, _
    ByVal vstrFileName As String) _
    As String

    Const cstrMethodName As String = "TransformErrorResponse"
    On Error GoTo TransformErrorResponseVbErr
    
    Dim xmlDoc As DOMDocument40
    Set xmlDoc = New DOMDocument40
    xmlDoc.setProperty "NewParser", True
    xmlDoc.async = False
    xmlDoc.loadXML vxmlResponse
    
    If xmlDoc.parseError.errorCode <> 0 Then
        xmlParseError xmlDoc.parseError
    End If
    
    If Not xmlDoc.selectSingleNode("RESPONSE[@TYPE='APPERR']") Is Nothing _
    Or Not xmlDoc.selectSingleNode("RESPONSE[@TYPE='SYSERR']") Is Nothing Then
        xmlTransform xmlDoc, vstrFileName
        TransformErrorResponse = xmlDoc.xml
    Else
        TransformErrorResponse = vxmlResponse
    End If
    
TransformErrorResponseVbErr:

    Set xmlDoc = Nothing
    
    CheckError cstrMethodName

End Function

' PSC 07/11/2006 EP2_41 - Start
Private Function CalculateKFIFees(ByVal vstrRequest As String) As String
    Const cstrMethodName As String = "CalculateKFIFees"
    Dim omigaObj As Object
    Dim xmlInDoc As DOMDocument40
    Dim xmlResponseDoc As DOMDocument40
    Dim objContext As ObjectContext
    Dim strResponse As String
    
    On Error GoTo CalculateKFIFeesVbErr
    
    Set xmlInDoc = New DOMDocument40
    xmlInDoc.setProperty "NewParser", True
    xmlInDoc.async = False
    xmlInDoc.loadXML vstrRequest
    
    Set xmlResponseDoc = New DOMDocument40
    xmlResponseDoc.setProperty "NewParser", True
    xmlResponseDoc.async = False

    
    If xmlInDoc.parseError.errorCode <> 0 Then
        xmlParseError xmlInDoc.parseError
    End If
    
    xmlTransform xmlInDoc, "transformCalculateKFIFeesRequest.xslt"
    
    Set objContext = GetObjectContext()
    If objContext Is Nothing Then
        Set omigaObj = CreateObject("omCK.CreateKFIBO")
    Else
        Set omigaObj = objContext.CreateInstance("omCK.CreateKFIBO")
    End If
    
    strResponse = omigaObj.CalculateKFIFees(xmlInDoc.xml)
    
    xmlResponseDoc.loadXML (strResponse)
    
    If Not xmlResponseDoc.selectSingleNode("RESPONSE[@TYPE='SUCCESS']") Is Nothing Then
        xmlTransform xmlResponseDoc, "transformCalculateKFIFeesResponse.xslt"
    End If

    CalculateKFIFees = xmlResponseDoc.xml
    
CalculateKFIFeesVbErr:

    Set omigaObj = Nothing
    Set xmlInDoc = Nothing
    
    CheckError cstrMethodName
End Function
' PSC 07/11/2006 EP2_41 - End

'EP2_421 - Web specific combo items
Private Function GetComboList(ByVal vstrRequest As String) As String

Const cstrMethodName As String = "GetComboList"
Const cstrRequestXSLT As String = "transformGetComboListRequest"
Const cstrResponseXSLT As String = "transformGetComboListResponse.xslt"
Const cstrWSComboListXML As String = "GetComboListWS.xml"
Const cstrResponsePath As String = "/RESPONSE"
Const cstrWSComboListNodePath As String = "/COMBOLISTWS"

Dim strCrudResponse As String
Dim strFileSpec As String
Dim strResult As String
Dim xmlCrudResponse As MSXML2.DOMDocument
Dim xmlComboList As MSXML2.DOMDocument
Dim objNode As IXMLDOMNode

    'Initialise
    On Error GoTo GetComboListExit
    strResult = ""
    Set xmlCrudResponse = New MSXML2.DOMDocument40
    xmlCrudResponse.async = False
    
    'Call CRUD to get back the unfiltered set of combos
    strCrudResponse = CrudBoInterface(vstrRequest, cstrRequestXSLT)
        
    If xmlCrudResponse.loadXML(strCrudResponse) Then
        
        'Load the list of combo items to include
        strFileSpec = App.Path & "\" & cstrWSComboListXML
        strFileSpec = Replace(strFileSpec, "DLL", "XML", 1, 1, vbTextCompare)
        
        Set xmlComboList = New MSXML2.DOMDocument40
        xmlComboList.async = False
        If xmlComboList.Load(strFileSpec) Then
        
            'Append the list of included combo items to the crud response.
            'This allows the xslt to have access to all the relevant data
            Set objNode = xmlCrudResponse.selectSingleNode(cstrResponsePath)
            If Not objNode Is Nothing Then
                objNode.appendChild xmlComboList.selectSingleNode(cstrWSComboListNodePath)
            End If
            
            'Transform the document filtering out all combo items that aren't in the included list
            xmlTransform xmlCrudResponse, cstrResponseXSLT
            strResult = xmlCrudResponse.xml
            
        End If
        
    End If
            
GetComboListExit:

    Set xmlCrudResponse = Nothing
    Set xmlComboList = Nothing
    Set objNode = Nothing
    GetComboList = strResult
    
    CheckError cstrMethodName
End Function

'EP2_2058 GHun
Private Function AssignCaseToPackager(ByVal vstrRequest As String) As String
    Const cstrFunctionName  As String = "AssignCaseToPackager"
    Dim strResponse         As String
    Dim xmlDoc              As FreeThreadedDOMDocument40
    'EP2_2182 GHun
    Dim strRequest          As String
    Dim objCRUD             As Object
    'EP2_2182 End

On Error GoTo AssignCaseToPackagerErr

    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.setProperty "NewParser", True
    xmlDoc.async = False

    'EP2_2182 GHun
    If xmlDoc.loadXML(vstrRequest) Then
        
        xmlTransform xmlDoc, "AssignCaseToPackagerRequest.xslt"
        
        strRequest = xmlDoc.xml

        Set objCRUD = GetObjectContext.CreateInstance("omCRUD.omCRUDBO")
        strResponse = objCRUD.omRequest(strRequest)
    'EP2_2182 End
        If xmlDoc.loadXML(strResponse) Then
            If xmlDoc.documentElement.getAttribute("TYPE") = "SUCCESS" Then
                CallHandleInterfaceResponse strRequest, "FE", "FulfilmentEmailMessageType", "ACTP" 'EP2_2058 GHun
            End If
        End If
        
        AssignCaseToPackager = strResponse
    'EP2_2182 GHun
    Else
        xmlParseError xmlDoc.parseError
    End If
    'EP2_2182 End

AssignCaseToPackagerErr:
    Set objCRUD = Nothing
    Set xmlDoc = Nothing
    CheckError cstrFunctionName
End Function
'EP2_2058 End

'EP2_2058 GHun
Private Sub CallHandleInterfaceResponse(ByVal vstrRequest As String, _
                                        ByVal vstrInterfaceTypeValidationType As String, _
                                        ByVal vstrMessageTypeCombo As String, _
                                        ByVal vstrMessageTypeValidationType As String)
    
    Const cstrFunctionName As String = "CallHandleInterfaceResponse"
    Dim objOmTMBO           As Object
    Dim objCRUD             As Object
    Dim xmlDoc              As FreeThreadedDOMDocument40
    Dim xmlRequest          As IXMLDOMElement
    Dim xmlElem             As IXMLDOMElement
    Dim xmlInRequest        As IXMLDOMElement
    Dim strResponse         As String
    Dim strValueId          As String
    Dim strMessageType      As String
    Dim strApplicationNum   As String
    
On Error GoTo CallHandleInterfaceResponseErr

    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.setProperty "NewParser", True
    xmlDoc.async = False
    
    If xmlDoc.loadXML(vstrRequest) Then 'EP2_2182 GHun
        Set xmlInRequest = xmlDoc.documentElement
        
        Set xmlRequest = xmlDoc.createElement("REQUEST")
        xmlRequest.setAttribute "USERID", xmlInRequest.getAttribute("USERID")
        xmlRequest.setAttribute "UNITID", xmlInRequest.getAttribute("UNITID")
        xmlRequest.setAttribute "CHANNELID", xmlInRequest.getAttribute("CHANNELID")
        xmlRequest.setAttribute "ENTITY_REF", "COMBOVALIDATION"
        xmlRequest.setAttribute "CRUD_OP", "READ"
        
        Set xmlElem = xmlDoc.createElement("COMBOVALIDATION")
        xmlElem.setAttribute "GROUPNAME", "InterfaceType"
        xmlElem.setAttribute "VALIDATIONTYPE", vstrInterfaceTypeValidationType
        xmlRequest.appendChild xmlElem
        
        'This may need changing when this function is used for other messages
        Set xmlElem = xmlInRequest.firstChild
        If Not xmlElem Is Nothing Then
            strApplicationNum = xmlElem.getAttribute("APPLICATIONNUMBER")
        End If
        
        Set xmlElem = xmlDoc.createElement("COMBOVALIDATION")
        xmlElem.setAttribute "GROUPNAME", vstrMessageTypeCombo
        xmlElem.setAttribute "VALIDATIONTYPE", vstrMessageTypeValidationType
        xmlRequest.appendChild xmlElem
            
        Set objCRUD = GetObjectContext.CreateInstance("omCRUD.omCRUDBO")
        strResponse = objCRUD.omRequest(xmlRequest.xml)
        
        xmlDoc.loadXML strResponse
        Set xmlElem = xmlDoc.selectSingleNode("RESPONSE/COMBOVALIDATION[@GROUPNAME='InterfaceType']")
        strValueId = xmlElem.getAttribute("VALUEID")
        
        Set xmlElem = xmlDoc.selectSingleNode("RESPONSE/COMBOVALIDATION[@GROUPNAME='" & vstrMessageTypeCombo & "']")
        'Only continue if a MessageType has been found
        If Not xmlElem Is Nothing Then
            strMessageType = xmlElem.getAttribute("VALUEID")
                
            Set xmlRequest = xmlDoc.createElement("REQUEST")
            xmlRequest.setAttribute "USERID", xmlInRequest.getAttribute("USERID")
            xmlRequest.setAttribute "UNITID", xmlInRequest.getAttribute("UNITID")
            xmlRequest.setAttribute "CHANNELID", xmlInRequest.getAttribute("CHANNELID")
            xmlRequest.setAttribute "USERAUTHORITYLEVEL", xmlInRequest.getAttribute("USERAUTHORITYLEVEL")
            xmlRequest.setAttribute "OPERATION", "HANDLEINTERFACERESPONSE"
            
            Set xmlElem = xmlDoc.createElement("APPLICATION")
            xmlElem.setAttribute "APPLICATIONNUMBER", strApplicationNum
            xmlElem.setAttribute "APPLICATIONFACTFINDNUMBER", "1"
            xmlRequest.appendChild xmlElem
            
            Set xmlElem = xmlDoc.createElement("INTERFACE")
            xmlElem.setAttribute "INTERFACETYPE", strValueId
            xmlElem.setAttribute "MESSAGETYPE", strMessageType
            xmlElem.setAttribute "CREATETASKFLAG", "1"
            xmlRequest.appendChild xmlElem
            
            Set objOmTMBO = GetObjectContext.CreateInstance("omTM.omTMBO")
            strResponse = objOmTMBO.omTmRequest(xmlRequest.xml)
            'There is no point in checking the omTMBO response as failures should be ignored
        End If
    Else
        xmlParseError xmlDoc.parseError 'EP2_2182 GHun
    End If
    
CallHandleInterfaceResponseErr:
    Set objCRUD = Nothing
    Set objOmTMBO = Nothing
    Set xmlDoc = Nothing
    Set xmlRequest = Nothing
    Set xmlInRequest = Nothing
    Set xmlElem = Nothing
    CheckError cstrFunctionName
End Sub
'EP2_1903 End
