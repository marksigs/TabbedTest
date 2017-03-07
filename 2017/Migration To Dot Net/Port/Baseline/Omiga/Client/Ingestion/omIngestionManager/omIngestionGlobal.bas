Attribute VB_Name = "omIngestionGlobal"
'-----------------------------------------------------------------------------
'Prog   Date        Description
'IK     18/10/2005  created for Project MARS
'IK     25/11/2005  MAR695 gstrComponentId & gstrComponentResponse now private member variables
'GHun   24/01/2006  MAR1084 Added errGetOmigaErrorNumber
'IK     28/03/2006  EP304 - Add TaskName to config for use by KYC task
'IK     11/10/2006  EP2_10 - address targetting
'IK     27/11/2006  EP2_202 add static data for address targeting
'IK     05/12/2006  EP2_287 'Shopping List' enhancements
'IK     08/01/2007  EP2_352 remote underwriter,
'IK     22/01/2007  EP2_920 simplify error messages
'IK     15/01/2007  EP2_983 re-fix error handling
'IK     21/03/2007  EP2_1062 cache App Form Template data

Option Explicit

Public gxmldocConfig As DOMDocument40
Public gstrComponentId As String
Public gstrComponentResponse As String
' EP2_202
Public gxmldocStaticData As DOMDocument40
' EP2_1062
Public gxmldocAppFormHostTemplate As DOMDocument40

' EP2_202
Private Sub LoadStaticData()

    Const cstrMethodName As String = "LoadStaticData"
    On Error GoTo LoadStaticDataExit

    Dim xmlRequest As DOMDocument40
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMElement
    
    Dim objCRUD As Object
    
    Set xmlRequest = New DOMDocument40
    xmlRequest.async = False
    xmlRequest.setProperty "NewParser", True
    xmlRequest.validateOnParse = False
    
    Set xmlNode = xmlRequest.appendChild(xmlRequest.createElement("REQUEST"))
    
    Set xmlElem = xmlRequest.createElement("OPERATION")
    xmlElem.setAttribute "CRUD_OP", "READ"
    
    For Each xmlNode In gxmldocConfig.documentElement.selectNodes("GLOBALPARAMETER")
        xmlElem.appendChild xmlNode.cloneNode(True)
    Next
    
    If xmlElem.hasChildNodes Then
        xmlRequest.documentElement.appendChild xmlElem
    End If
    
    Set xmlElem = xmlRequest.createElement("OPERATION")
    xmlElem.setAttribute "CRUD_OP", "READ"
    
    For Each xmlNode In gxmldocConfig.documentElement.selectNodes("COMBOGROUP")
        xmlElem.appendChild xmlNode.cloneNode(True)
    Next
    
    If xmlElem.hasChildNodes Then
        xmlRequest.documentElement.appendChild xmlElem
    End If
    
    If xmlRequest.documentElement.hasChildNodes Then
    
        Set gxmldocStaticData = New DOMDocument40
        gxmldocStaticData.async = False
        gxmldocStaticData.setProperty "NewParser", True
        gxmldocStaticData.validateOnParse = False
    
        Set objCRUD = GetObjectContext.CreateInstance("omCRUD.omCRUDBO")
        gxmldocStaticData.loadXML objCRUD.OmRequest(xmlRequest.xml)
        Set objCRUD = Nothing
    
        If gxmldocStaticData.parseError.errorCode <> 0 Then
            xmlParseError gxmldocConfig.parseError
        End If
    
    End If
    
LoadStaticDataExit:

    Set xmlNode = Nothing
    Set xmlElem = Nothing
    Set xmlRequest = Nothing
    
    If Err.Number <> 0 Then
        If Err.Source = cstrMethodName Then
            Err.Raise Err.Number, Err.Source, Err.Description
        Else
            Err.Raise Err.Number, cstrMethodName & "." & Err.Source, Err.Description
        End If
    End If
    
End Sub

Public Function GetWebChannelId() As String
    GetWebChannelId = GetGlobalAmount("WebChannelId", "2")
End Function

Public Function GetProductCascadeDecision() As String
    GetProductCascadeDecision = GetComboValueId("RiskAssessmentDecision", "AC", "15")
End Function

Public Function GetApplicantCustomerRoleType() As String
    GetApplicantCustomerRoleType = GetComboValueId("CustomerRoleType", "A", "1")
End Function

Public Function GetDefaultValuationType() As String
    GetDefaultValuationType = GetComboValueId("ValuationType", "SC1", "1")
End Function

Public Function GetNoAuthorityUserAccessType() As String
    GetNoAuthorityUserAccessType = GetComboValueId("UserAccessType", "10", "10")
End Function

Public Function GetDefaultPropertyLocation() As String
    GetDefaultPropertyLocation = GetComboValueId("PropertyLocation", "E", "10")
End Function

Public Function GetGenderValidationType(ByVal vstrValueId As String) As String
    GetGenderValidationType = GetComboValidationType("Sex", vstrValueId)
End Function

Public Function GetGlobalAmount( _
    ByVal vstrGlobalName As String, _
    ByVal vstrDefaultValue As String) _
    As String
    
    Dim xmlElem As IXMLDOMElement
    
    Set xmlElem = _
        gxmldocStaticData.selectSingleNode("RESPONSE/OPERATION/RESPONSE/GLOBALPARAMETER[@NAME='" & vstrGlobalName & "'][@AMOUNT]")
        
    If Not xmlElem Is Nothing Then
        GetGlobalAmount = xmlElem.getAttribute("AMOUNT")
    Else
        GetGlobalAmount = vstrDefaultValue
    End If

    Set xmlElem = Nothing

End Function

'EP2_352
Public Function GetGlobalBool( _
    ByVal vstrGlobalName As String, _
    ByVal vblnDefaultValue As Boolean) _
    As Boolean
    
    Dim xmlElem As IXMLDOMElement
    
    GetGlobalBool = vblnDefaultValue
    
    Set xmlElem = _
        gxmldocStaticData.selectSingleNode("RESPONSE/OPERATION/RESPONSE/GLOBALPARAMETER[@NAME='" & vstrGlobalName & "'][@BOOLEAN]")
        
    If Not xmlElem Is Nothing Then
        If xmlElem.getAttribute("BOOLEAN") = 1 Then
            GetGlobalBool = True
        Else
            GetGlobalBool = False
        End If
    End If

    Set xmlElem = Nothing

End Function

Public Function GetGlobalString( _
    ByVal vstrGlobalName As String, _
    ByVal vstrDefaultValue As String) _
    As String
    
    Dim xmlElem As IXMLDOMElement
    
    Set xmlElem = _
        gxmldocStaticData.selectSingleNode("RESPONSE/OPERATION/RESPONSE/GLOBALPARAMETER[@NAME='" & vstrGlobalName & "'][@STRING]")
        
    If Not xmlElem Is Nothing Then
        GetGlobalString = xmlElem.getAttribute("STRING")
    Else
        GetGlobalString = vstrDefaultValue
    End If

    Set xmlElem = Nothing

End Function
'EP2_352_ends

Public Function GetComboValueId( _
    ByVal vstrGroupName As String, _
    ByVal vstrValidationType As String, _
    ByVal vstrDefaultValue As String) _
    As String
    
    Dim xmlElem As IXMLDOMElement
    
    Set xmlElem = _
        gxmldocStaticData.selectSingleNode("RESPONSE/OPERATION/RESPONSE/COMBOGROUP[@GROUPNAME='" & vstrGroupName & "']/COMBOVALUE[COMBOVALIDATION[@VALIDATIONTYPE='" & vstrValidationType & "']]")
        
    If Not xmlElem Is Nothing Then
        GetComboValueId = xmlElem.getAttribute("VALUEID")
    Else
        GetComboValueId = vstrDefaultValue
    End If

    Set xmlElem = Nothing

End Function

Public Function GetComboValidationType( _
    ByVal vstrGroupName As String, _
    ByVal vstrValueId As String) _
    As String
    
    Dim xmlElem As IXMLDOMElement
    
    Set xmlElem = _
        gxmldocStaticData.selectSingleNode("RESPONSE/OPERATION/RESPONSE/COMBOGROUP[@GROUPNAME='" & vstrGroupName & "']/COMBOVALUE[@VALUEID='" & vstrValueId & "']/COMBOVALIDATION")
        
    If Not xmlElem Is Nothing Then
        GetComboValidationType = xmlElem.getAttribute("VALIDATIONTYPE")
    End If

    Set xmlElem = Nothing

End Function

Public Function IsComboValidationType( _
    ByVal vstrGroupName As String, _
    ByVal vstrValueId As String, _
    ByVal vstrValidationType As String) _
    As Boolean
    
    Dim xmlElem As IXMLDOMElement
    
    Set xmlElem = _
        gxmldocStaticData.selectSingleNode("RESPONSE/OPERATION/RESPONSE/COMBOGROUP[@GROUPNAME='" & vstrGroupName & "']/COMBOVALUE[@VALUEID='" & vstrValueId & "']/COMBOVALIDATION[@VALIDATIONTYPE='" & vstrValidationType & "']")
        
    If Not xmlElem Is Nothing Then
        IsComboValidationType = True
    End If

    Set xmlElem = Nothing

End Function

' EP2_202_ends

Public Sub LoadConfigFile()

    Const cstrMethodName As String = "LoadConfigFile"
    On Error GoTo LoadConfigFileExit
    
    Dim strFileSpec As String
    
    ' pick up XML map from "...\Omiga 4\XML" directory
    ' Only do the subsitution once to change DLL -> XML
    strFileSpec = App.Path & "\ingestionConfig.xml"
    strFileSpec = Replace(strFileSpec, "DLL", "XML", 1, 1, vbTextCompare)
    
    Set gxmldocConfig = New DOMDocument40
    gxmldocConfig.async = False
    gxmldocConfig.setProperty "NewParser", True
    gxmldocConfig.validateOnParse = False
    gxmldocConfig.Load strFileSpec
    
    If gxmldocConfig.parseError.errorCode <> 0 Then
        xmlParseError gxmldocConfig.parseError
    End If

    ' EP2_202
    Call LoadStaticData
    
LoadConfigFileExit:
    
    If Err.Number <> 0 Then
        If Err.Source = cstrMethodName Then
            Err.Raise Err.Number, Err.Source, Err.Description
        Else
            Err.Raise Err.Number, cstrMethodName & "." & Err.Source, Err.Description
        End If
    End If

End Sub

Public Sub xmlParseError(ByVal objParseError As IXMLDOMParseError)
    
    Dim strErrDesc As String    ' formatted parser error
    strErrDesc = _
        "XML parser error - " & vbCr & _
        "Reason: " & objParseError.reason & vbCr & _
        "Error code: " & Str$(objParseError.errorCode) & vbCr & _
        "Line no.: " & Str$(objParseError.Line) & vbCr & _
        "Character: " & Str$(objParseError.linepos) & vbCr & _
        "Source text: " & objParseError.srcText
        
        Err.Raise oeXMLParserError, "xmlParseError", strErrDesc
            
End Sub

Public Sub CheckError(ByVal vstrMethodName As String)
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

Public Function ChainErrorSource(ByVal vstrMethodName As String) As String
    
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

Public Sub ChainResponseDescription(ByVal vxmlResponseDoc As DOMDocument40, ByVal vstrDesc As String)
    If Not vxmlResponseDoc.selectSingleNode("RESPONSE/ERROR/DESCRIPTION") Is Nothing Then
        vxmlResponseDoc.selectSingleNode("RESPONSE/ERROR/DESCRIPTION").Text = vstrDesc & ", " & vxmlResponseDoc.selectSingleNode("RESPONSE/ERROR/SOURCE").Text
    End If
End Sub

Public Function FormatError( _
    ByVal vstrXmlIn As String, _
    ByVal vstrComponentId As String, _
    ByVal vstrComponentResponse) As String
    
    Dim xmlErrDoc As DOMDocument40
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    
    Set xmlErrDoc = New DOMDocument40
    xmlErrDoc.setProperty "NewParser", True
    xmlErrDoc.async = False
    
    Set xmlElem = xmlErrDoc.createElement("RESPONSE")
    
    If InStr(Err.Source, App.EXEName) <> 0 Then
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
    
    ' EP2_1062
    ' remove sensitive data from error response
    If Not gxmldocConfig.selectSingleNode("omIngestion[@errorDetail]") Is Nothing Then
    
        Set xmlElem = xmlErrDoc.createElement("REQUEST")
        xmlElem.Text = vstrXmlIn
        xmlNode.appendChild xmlElem
    
        If Len(vstrComponentResponse) <> 0 Then
            Set xmlElem = xmlErrDoc.createElement("COMPONENT_ID")
            xmlElem.Text = vstrComponentId
            xmlNode.appendChild xmlElem
            Set xmlElem = xmlErrDoc.createElement("COMPONENT_RESPONSE")
            xmlElem.Text = vstrComponentResponse
            xmlNode.appendChild xmlElem
        End If
        
    End If
    
    FormatError = xmlErrDoc.xml
    
    Set xmlNode = Nothing
    Set xmlElem = Nothing
    Set xmlErrDoc = Nothing

End Function

Public Function GetApplicationData(ByVal vxmlRequestNode As IXMLDOMNode) As String
    
    Const cstrMethodName As String = "GetApplicationData"
    On Error GoTo GetApplicationDataExit
    
    Dim crudObj As Object
    
    Dim strApplicationNumber As String, _
        strApplicationFactFindNumber As String, _
        strRequest As String
    
    strApplicationNumber = vxmlRequestNode.selectSingleNode("APPLICATION/@APPLICATIONNUMBER").Text
    If vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/@APPLICATIONFACTFINDNUMBER") Is Nothing Then
        strApplicationFactFindNumber = "1"
    Else
        strApplicationFactFindNumber = vxmlRequestNode.selectSingleNode("APPLICATION/APPLICATIONFACTFIND/@APPLICATIONFACTFINDNUMBER").Text
    End If
    
    strRequest = _
        "<REQUEST CRUD_OP='READ' SCHEMA_NAME='WebServiceSchema' ENTITY_REF='AIPINGESTION'>" & _
        "<APPLICATION APPLICATIONNUMBER='" & strApplicationNumber & _
        "' APPLICATIONFACTFINDNUMBER='" & strApplicationFactFindNumber & "'/></REQUEST>"

    Set crudObj = GetObjectContext.CreateInstance("omCRUD.omCRUDBO")
    GetApplicationData = crudObj.OmRequest(strRequest)
    
GetApplicationDataExit:
    
    Set crudObj = Nothing
    
    CheckError cstrMethodName

End Function

Public Sub GetCaseActivity(ByVal vxmlRequestNode As IXMLDOMNode, ByVal vxmlResponseDoc As DOMDocument40)
    
    Const cstrMethodName As String = "GetCaseActivity"
    On Error GoTo GetCaseActivityExit
    
    Dim xmlDoc As DOMDocument40
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMNode
    
    Dim crudObj As Object
    
    Set xmlDoc = New DOMDocument40
    xmlDoc.async = False
    xmlDoc.setProperty "NewParser", True
    
    Set xmlElem = xmlDoc.createElement("REQUEST")
    xmlElem.setAttribute "CRUD_OP", "READ"
    Set xmlNode = xmlDoc.appendChild(xmlElem)
    
    Set xmlElem = xmlDoc.createElement("APPLICATIONPRIORITY")
    xmlElem.setAttribute "APPLICATIONNUMBER", vxmlRequestNode.selectSingleNode("APPLICATION/@APPLICATIONNUMBER").Text
    xmlNode.appendChild xmlElem
    
    Set xmlElem = xmlDoc.createElement("CASEACTIVITY")
    xmlElem.setAttribute "CASEID", vxmlRequestNode.selectSingleNode("APPLICATION/@APPLICATIONNUMBER").Text
    Set xmlNode = xmlNode.appendChild(xmlElem)
    
    Set xmlElem = xmlDoc.createElement("CASESTAGE")
    Set xmlNode = xmlNode.appendChild(xmlElem)
    
    Set xmlElem = xmlDoc.createElement("CASETASK")
    xmlNode.appendChild xmlElem
    
    Set crudObj = GetObjectContext.CreateInstance("omCRUD.omCRUDBO")
    vxmlResponseDoc.loadXML crudObj.OmRequest(xmlDoc.xml)
    Set crudObj = Nothing
    
    If vxmlResponseDoc.parseError.errorCode <> 0 Then
        xmlParseError vxmlResponseDoc.parseError
    End If
    
    If vxmlResponseDoc.selectSingleNode("RESPONSE[@TYPE='SUCCESS']") Is Nothing Then
        Err.Raise oeUnspecifiedError, cstrMethodName, vxmlResponseDoc.xml
    End If
    
    If vxmlResponseDoc.selectSingleNode("RESPONSE/APPLICATIONPRIORITY") Is Nothing Then
        Err.Raise oeXMLMissingElement, cstrMethodName, "error retrieving APPLICATIONPRIORITY details"
    End If
    
    If vxmlResponseDoc.selectSingleNode("RESPONSE/CASEACTIVITY/CASESTAGE") Is Nothing Then
        Err.Raise oeXMLMissingElement, cstrMethodName, "error retrieving CASEACTIVITY details"
    End If
    
GetCaseActivityExit:
    
    CheckError cstrMethodName

End Sub

Public Function OutstandingTaskExists( _
    ByVal vxmlCaseStage As IXMLDOMNode, _
    ByVal vstrTaskId As String) _
    As Boolean
    
    Const cstrMethodName As String = "OutstandingTaskExists"
    On Error GoTo OutstandingTaskExistsExit
    
    Dim strXpath As String
        
    strXpath = "CASETASK[@TASKID='" & vstrTaskId & "'][@TASKSTATUS='10' or @TASKSTATUS='20']"
    
    If vxmlCaseStage.selectNodes(strXpath).length > 0 Then
        OutstandingTaskExists = True
    End If
    
OutstandingTaskExistsExit:
    
    CheckError cstrMethodName

End Function

Public Function GetOutstandingTask( _
    ByVal vxmlCaseStage As IXMLDOMNode, _
    ByVal vstrTaskId As String) _
    As IXMLDOMNode
    
    Const cstrMethodName As String = "GetOutstandingTask"
    On Error GoTo GetOutstandingTaskExit
    
    Dim strXpath As String
        
    strXpath = "CASETASK[@TASKID='" & vstrTaskId & "'][@TASKSTATUS='10' or @TASKSTATUS='20']"
    
    Set GetOutstandingTask = vxmlCaseStage.selectSingleNode(strXpath)
    
GetOutstandingTaskExit:
    
    CheckError cstrMethodName

End Function


Public Function GetTaskId(ByVal vstrCreditCheckOp As String) As String
    
    Const cstrMethodName As String = "GetTaskId"
    On Error GoTo GetTaskIdExit
    
    If gxmldocConfig.selectSingleNode("omIngestion/operation[@name='" & vstrCreditCheckOp & "'][@task]") Is Nothing Then
        Err.Raise oeXMLMissingAttribute, cstrMethodName, "no taskid in ingestionConfig.xml for credit check operation " & vstrCreditCheckOp
    Else
        GetTaskId = gxmldocConfig.selectSingleNode("omIngestion/operation[@name='" & vstrCreditCheckOp & "'][@task]").Attributes.getNamedItem("task").Text
    End If
    
GetTaskIdExit:
    
    CheckError cstrMethodName

End Function

'IK_EP2_10_09/10/2006
Public Function GetTaskInterface(ByVal vstrCreditCheckOp As String) As String
    
    Const cstrMethodName As String = "GetTaskInterface"
    On Error GoTo GetTaskInterfaceExit
    
    Dim objCRUD As Object
    Dim xmlDoc As DOMDocument40
    
    Dim strRequest As String
    
    strRequest = _
        "<REQUEST CRUD_OP='READ'><TASK TASKID='" & _
        GetTaskId(vstrCreditCheckOp) & _
        "'/></REQUEST>"
            
    Set xmlDoc = New DOMDocument40
    xmlDoc.async = False
    xmlDoc.setProperty "NewParser", True
    xmlDoc.loadXML strRequest
    
    Set objCRUD = GetObjectContext.CreateInstance("omCRUD.omCRUDBO")
    xmlDoc.loadXML objCRUD.OmRequest(xmlDoc.xml)
    Set objCRUD = Nothing
    
    If Not xmlDoc.selectSingleNode("RESPONSE[@TYPE='SUCCESS']/TASK[@INTERFACE]") Is Nothing Then
        GetTaskInterface = xmlDoc.selectSingleNode("RESPONSE/TASK/@INTERFACE").Text
        Set xmlDoc = Nothing
    Else
        Set xmlDoc = Nothing
        Err.Raise oeXMLMissingAttribute, cstrMethodName, "no task interface for credit check operation " & vstrCreditCheckOp
    End If
    
GetTaskInterfaceExit:
    
    CheckError cstrMethodName

End Function


'IK_EP304_28/03/2006
'get task name (via omCRUD) & append to config file
Public Function GetTaskname(ByVal vstrCreditCheckOp As String) As String
    
    Const cstrMethodName As String = "GetTaskName"
    On Error GoTo GetTaskNameExit
    
    Dim xmlDoc As DOMDocument40
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMElement
    Dim xmlConfigNode As IXMLDOMElement
    
    Dim objCRUD As Object
    
    Set xmlConfigNode = gxmldocConfig.selectSingleNode("omIngestion/operation[@name='" & vstrCreditCheckOp & "']")
    
    If Not xmlConfigNode Is Nothing Then
        
        If xmlConfigNode.Attributes.getNamedItem("taskName") Is Nothing Then
    
            Set xmlDoc = New DOMDocument40
            xmlDoc.async = False
            xmlDoc.setProperty "NewParser", True
    
            Set xmlElem = xmlDoc.createElement("REQUEST")
            xmlElem.setAttribute "CRUD_OP", "READ"
            Set xmlNode = xmlDoc.appendChild(xmlElem)
            
            Set xmlElem = xmlDoc.createElement("TASK")
            xmlElem.setAttribute "TASKID", xmlConfigNode.getAttribute("task")
            xmlNode.appendChild xmlElem
            
            Set objCRUD = GetObjectContext.CreateInstance("omCRUD.omCRUDBO")
            xmlDoc.loadXML objCRUD.OmRequest(xmlDoc.xml)
    
            If xmlDoc.selectSingleNode("RESPONSE/TASK[@TASKNAME]") Is Nothing Then
                Err.Raise oeXMLMissingAttribute, cstrMethodName, "error reading TASKNAME for credit check operation " & vstrCreditCheckOp
            End If
            
            xmlConfigNode.setAttribute "taskName", xmlDoc.selectSingleNode("RESPONSE/TASK/@TASKNAME").Text
            
        End If
        
        GetTaskname = xmlConfigNode.getAttribute("taskName")
        
    End If
    
GetTaskNameExit:
    
    Set objCRUD = Nothing
    Set xmlNode = Nothing
    Set xmlElem = Nothing
    Set xmlDoc = Nothing
    Set xmlConfigNode = Nothing
    
    CheckError cstrMethodName

End Function
'IK_EP304_28/03/2006_ends


Public Sub xmlTransform(ByVal xmlDoc As DOMDocument40, ByVal vstrFileName)

    Const cstrMethodName As String = "xmlTransform"
    
    Dim xslDoc As DOMDocument40
    
    Dim strFileSpec As String
    
    strFileSpec = App.Path & "\" & vstrFileName
    strFileSpec = Replace(strFileSpec, "DLL", "XML", 1, 1, vbTextCompare)
    
    Set xslDoc = New DOMDocument40
    xslDoc.async = False
    xslDoc.Load strFileSpec
    
    If xslDoc.parseError.errorCode <> 0 Then
        Err.Raise _
            oeXMLParserError, _
            cstrMethodName, _
                "error loading xslt document: " & _
                strFileSpec & _
                " " & xslDoc.parseError.reason
    End If
    
    xmlDoc.loadXML xmlDoc.transformNode(xslDoc)
    
    If xslDoc.parseError.errorCode <> 0 Then
        Err.Raise _
            oeXMLParserError, _
            cstrMethodName, _
                "error in transformation: " & _
                strFileSpec & _
                " " & xslDoc.parseError.reason
    End If

End Sub

Public Function FormatNow() As String

    Dim dtNow As Date
    dtNow = Now()
    
    FormatNow = CStr(Day(dtNow)) & "/" & CStr(Month(dtNow)) & "/" & CStr(Year(dtNow)) & _
                " " & CStr(Hour(dtNow)) & ":" & CStr(Minute(dtNow)) & ":" & CStr(Second(dtNow))

End Function

'MAR1084 GHun
Public Function errGetOmigaErrorNumber(ByVal vlngErrorNo As Long) As Long
' header ----------------------------------------------------------------------------------
' description:  Converts an  error number to an omiga number. When errors are raised by VB
'               they get added to them the VB constant 'vbObjectError + 512' so to get the
'               Omiga4 error number orginally raised we need to subtract them from the err
'               number.
'               NOTE: System error numbers will end up much larger than Omiga4 error numbers
' pass:         Error number to calculate
' return:       Calculated Omiga4 error number
'------------------------------------------------------------------------------------------
    errGetOmigaErrorNumber = vlngErrorNo - vbObjectError - 512
End Function
'MAR1084 End

' EP2_10
Public Function IsAddressTargetResolver(ByVal vxmlRequestNode As IXMLDOMNode) As Boolean
    If Not vxmlRequestNode.selectSingleNode("TARGETINGDATA") Is Nothing Then
        IsAddressTargetResolver = True
    End If
End Function

'EP2_352
Public Function CrudBOCall(ByVal vstrCrudRequest As String) As String
    Dim objCRUD As Object
    Set objCRUD = GetObjectContext.CreateInstance("omCRUD.omCRUDBO")
    CrudBOCall = objCRUD.OmRequest(vstrCrudRequest)
    Set objCRUD = Nothing
End Function
'EP2_352_ends

'EP2_920
Public Sub LogOmigaError(ByVal vstrOmigaResponse As String)
    App.LogEvent vstrOmigaResponse, vbLogEventTypeInformation
End Sub
'EP2_920_ends

' EP2_1062
Public Sub LoadAppFormHostTemplateData()
    
    Const cstrMethodName As String = "LoadAppFormHostTemplateData"
        
    Dim strRequest As String
    
    If gxmldocAppFormHostTemplate Is Nothing Then
    
        strRequest = _
            "<REQUEST CRUD_OP='READ'><HOSTTEMPLATE HOSTTEMPLATEID='" & _
            GetGlobalString("WebApplicationFormDocumentId", "058") & _
            "'/></REQUEST>"
            
        Set gxmldocAppFormHostTemplate = New DOMDocument40
        gxmldocAppFormHostTemplate.async = False
        gxmldocAppFormHostTemplate.setProperty "NewParser", True
        gxmldocAppFormHostTemplate.validateOnParse = False
    
        gxmldocAppFormHostTemplate.loadXML CrudBOCall(strRequest)
        
    End If
    
    If gxmldocAppFormHostTemplate.selectSingleNode("RESPONSE[@TYPE='SUCCESS']") Is Nothing Then
        LogOmigaError gxmldocAppFormHostTemplate.xml
        Err.Raise oeUnspecifiedError, cstrMethodName, "error retrieving Application Form Template details"
    End If
    
    If gxmldocAppFormHostTemplate.selectSingleNode("RESPONSE/HOSTTEMPLATE[@HOSTTEMPLATEID]") Is Nothing Then
        LogOmigaError gxmldocAppFormHostTemplate.xml
        Err.Raise oeUnspecifiedError, cstrMethodName, "Application Form Template details missing HOSTTEMPLATEID"
    End If
    
    If gxmldocAppFormHostTemplate.selectSingleNode("RESPONSE/HOSTTEMPLATE[@DPSTEMPLATEID]") Is Nothing Then
        LogOmigaError gxmldocAppFormHostTemplate.xml
        Err.Raise oeUnspecifiedError, cstrMethodName, "Application Form Template details missing DPSTEMPLATEID"
    End If
    
    If gxmldocAppFormHostTemplate.selectSingleNode("RESPONSE/HOSTTEMPLATE[@PDMMETHOD]") Is Nothing Then
        LogOmigaError gxmldocAppFormHostTemplate.xml
        Err.Raise oeUnspecifiedError, cstrMethodName, "Application Form Template details missing PDMMETHOD"
    End If
    
    If gxmldocAppFormHostTemplate.selectSingleNode("RESPONSE/HOSTTEMPLATE[@URLPOSTIND = '1']") Is Nothing Then
        LogOmigaError gxmldocAppFormHostTemplate.xml
        Err.Raise oeUnspecifiedError, cstrMethodName, "Application Form Template invalid, URLPOSTIND value incorrect"
    End If

End Sub
' EP2_1062_ends

