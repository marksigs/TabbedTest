Attribute VB_Name = "Methods"
'Workfile:      Methods.bas
'Copyright:     Copyright © 2001 Marlborough Stirling

'Description:
'
'Dependencies:
'Issues:        Instancing:
'               MTSTransactionMode:
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'RF     24/08/01    Expanded code created by LD.
'RF     12/09/01    Allow for ODIConverter interface change.
'RF     17/09/01    Change to Converter Request header.
'AS     04/10/01    Converter request now conforms to Omiga IV Phase 2 standards.
'RF     10/12/01    Added ChargeType data node default value 1.
'DS     21/03/02    Various fixes for building search keys. SYS4306.
'DM     03/04/02    SYS4350 Implement SaveThirdpartyDetails
'DS     30/04/02    Use FreeThreadedDOMDocument.
'STB    21/05/02    SYS4609 Implement SaveMortgageProductDetails, GetUserMortgageAdminDetails and SaveUserMortgageAdminDetails.
'STB    09/07/02    SYS5101 Implemented PlexusHomeImpl_ProcessRateChange().
'DRC    30/07/02    SYS4945 Added an extra (optional) arguement to PlexusHomeImpl_create_PrimaryMortgageKey for Chargetype
'SG/JR  21/11/2002  SYS5765/SYSMCP1256 - Added PlexusHomeImpl_findByPrimaryKey_ThirdPartyKey.
'------------------------------------------------------------------------------------------
Option Explicit

Private Function SendRequestToConverter( _
    ByVal vnodeRequest As IXMLDOMNode, _
    Optional ByVal objODIConverter As ODICONVERTER.ODICONVERTER = Nothing _
    ) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
' exceptions:
'------------------------------------------------------------------------------------------
On Error GoTo SendRequestToConverterExit
    
    Const strFunctionName As String = "SendRequestToConverter"
     
     ' <VSA> Visual Studio Analyser Support
    #If USING_VSA Then
        Dim VSA As New vsa_shared: VSA.Initialize (TypeName(Me) & "." & strFunctionName)
    #End If

    Dim strConverterResponse As String
    Dim xmlDoc As New FreeThreadedDOMDocument
    
    If objODIConverter Is Nothing Then
        Dim objTempODIConverter As ODICONVERTER.ODICONVERTER
        Set objTempODIConverter = New ODICONVERTER.ODICONVERTER
        strConverterResponse = objTempODIConverter.Request(vnodeRequest.xml)
        Set objTempODIConverter = Nothing
    Else
        strConverterResponse = objODIConverter.Request(vnodeRequest.xml)
    End If
    
    ' return the response as a node
    
    Set xmlDoc = xmlLoad(strConverterResponse, strFunctionName)
    Set SendRequestToConverter = xmlDoc.documentElement
    
SendRequestToConverterExit:

    Set objTempODIConverter = Nothing

    ' <VSA> Visual Studio Analyser Support
    #If USING_VSA Then
        Set VSA = Nothing
    #End If
    
    errCheckError strFunctionName

End Function

Public Function PlexusHomeImpl_create_PrimaryCustomerKey( _
    ByVal vobjODITransformerState As ODITransformerState, _
    Optional ByVal objODIConverter As ODICONVERTER.ODICONVERTER = Nothing _
    ) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
' exceptions:
'------------------------------------------------------------------------------------------
On Error GoTo PlexusHomeImpl_create_PrimaryCustomerKeyExit
    
    Const strFunctionName As String = "PlexusHomeImpl_create_PrimaryCustomerKey"
    
    Dim nodeConverterRequest As IXMLDOMNode
        
    Set nodeConverterRequest = CreateConverterRequest( _
        vobjODITransformerState, 1, CreatePrimaryCustomerKey(""), _
        "create", "PlexusHomeImpl")
    
    Set PlexusHomeImpl_create_PrimaryCustomerKey = _
        SendRequestToConverter(nodeConverterRequest, objODIConverter)
    
PlexusHomeImpl_create_PrimaryCustomerKeyExit:

    Set nodeConverterRequest = Nothing

    errCheckError strFunctionName
    
End Function

Public Function CreatePrimaryCustomerKey( _
    ByVal vstrEntityNumber As String) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
' exceptions:
'------------------------------------------------------------------------------------------
On Error GoTo CreatePrimaryCustomerKeyExit
    
    Const strFunctionName As String = "CreatePrimaryCustomerKey"
        
    Dim xmlDoc As FreeThreadedDOMDocument
    Dim elemPrimaryCustomerKey As IXMLDOMElement
    Dim elemTemp As IXMLDOMElement
    Dim xmlAttr As IXMLDOMAttribute
     
    Set xmlDoc = New FreeThreadedDOMDocument
    Set elemPrimaryCustomerKey = xmlDoc.createElement("PRIMARYCUSTOMERKEY")
       
    Set elemTemp = xmlDoc.createElement("ENTITYNUMBER")
    xmlSetAttributeValue elemTemp, "DATA", vstrEntityNumber
    
    elemPrimaryCustomerKey.appendChild elemTemp
    
    Set CreatePrimaryCustomerKey = elemPrimaryCustomerKey
        
CreatePrimaryCustomerKeyExit:

    errCheckError strFunctionName

End Function

Private Function CreateCustomerInvolvementPattern( _
    ByVal vstrEntityNumber As String) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
' exceptions:
'------------------------------------------------------------------------------------------
On Error GoTo CreateCustomerInvolvementPatternExit
    
    Const strFunctionName As String = "CreateCustomerInvolvementPattern"
            
    Dim xmlDoc As FreeThreadedDOMDocument
    Dim elemCustomerInvolvementPattern, elemTargetKey As IXMLDOMElement
    Dim elemTemp As IXMLDOMElement
     
    Set xmlDoc = New FreeThreadedDOMDocument
    Set elemCustomerInvolvementPattern = _
        xmlDoc.createElement("CUSTOMERINVOLVEMENTPATTERN")
        
    Set elemTargetKey = xmlDoc.createElement("TARGETKEY")
        
    Set elemTemp = CreatePrimaryCustomerKey(vstrEntityNumber)
       
    elemCustomerInvolvementPattern.appendChild elemTargetKey
    elemTargetKey.appendChild elemTemp
    
    Set CreateCustomerInvolvementPattern = elemCustomerInvolvementPattern
        
CreateCustomerInvolvementPatternExit:

    errCheckError strFunctionName
    
End Function
Private Function CreateCustomerListPattern( _
    ByVal vstrEntityNumber As String) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
' exceptions:
'------------------------------------------------------------------------------------------
On Error GoTo CreateCustomerListPatternExit
    
    Const strFunctionName As String = "CreateCustomerListPattern"
            
    Dim xmlDoc As FreeThreadedDOMDocument
    Dim elemCustomerListPattern, elemTargetKey As IXMLDOMElement
    Dim elemTemp As IXMLDOMElement
     
    Set xmlDoc = New FreeThreadedDOMDocument
    Set elemCustomerListPattern = _
        xmlDoc.createElement("CUSTOMERLISTPATTERN")
        
    Set elemTargetKey = xmlDoc.createElement("TARGETKEY")
        
    Set elemTemp = CreateMortgageKey(vstrEntityNumber)
       
    elemCustomerListPattern.appendChild elemTargetKey
    elemTargetKey.appendChild elemTemp
    
    Set CreateCustomerListPattern = elemCustomerListPattern
        
CreateCustomerListPatternExit:

    errCheckError strFunctionName
    
End Function

Public Function PlexusHomeImpl_create_PrimaryMortgageKey( _
    ByVal vobjODITransformerState As ODITransformerState, _
    Optional ByVal objODIConverter As ODICONVERTER.ODICONVERTER = Nothing, _
    Optional ByVal strChargeType As String = "" _
    ) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
' pass:
' return:
' exceptions:
'------------------------------------------------------------------------------------------
On Error GoTo PlexusHomeImpl_create_PrimaryMortgageKeyExit
    
    Const strFunctionName As String = "PlexusHomeImpl_create_PrimaryMortgageKey"
            
    Dim nodeConverterRequest As IXMLDOMNode
    'AQR SYS4945 - DRC checking for a specific charge type
    Dim strKeyArg As String
    If Len(strChargeType) = 0 Then
        strKeyArg = ""
    Else
        strKeyArg = "." & strChargeType
    End If
    
    Set nodeConverterRequest = CreateConverterRequest( _
        vobjODITransformerState, 1, CreateMortgageKey(strKeyArg), _
        "create", "PlexusHomeImpl")
    'AQR SYS4945 - end
    Set PlexusHomeImpl_create_PrimaryMortgageKey = _
        SendRequestToConverter(nodeConverterRequest, objODIConverter)
    
PlexusHomeImpl_create_PrimaryMortgageKeyExit:

    Set nodeConverterRequest = Nothing

    errCheckError strFunctionName

End Function

Public Function CreateMortgageKey( _
    ByVal vstrAccountNumber As String) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
' pass:
'   vstrAccountNumber
'       Account number in following format: collateralNumber + "." + chargeType
' return:
' exceptions:
'------------------------------------------------------------------------------------------
On Error GoTo CreateMortgageKeyExit
    
    Const strFunctionName As String = "CreateMortgageKey"
            
    Dim xmlDoc As FreeThreadedDOMDocument
    Dim elemMortgagekey As IXMLDOMElement
    Dim elemTemp As IXMLDOMElement
    Dim elemRealEstateKey, elemRealEstateKey2  As IXMLDOMElement
    Dim intSeparatorPosition As Integer
    Dim intAccountNumberLength As Integer
    
    intAccountNumberLength = Len(vstrAccountNumber)
    
    If intAccountNumberLength > 0 Then
        intSeparatorPosition = InStr(1, vstrAccountNumber, ".", vbTextCompare)
        If intSeparatorPosition < 1 Then
            errThrowError strFunctionName, oeInvalidParameter, _
                "Invalid Account Number format: " & vstrAccountNumber
        End If
    End If
     
    Set xmlDoc = New FreeThreadedDOMDocument
    Set elemMortgagekey = xmlDoc.createElement("MORTGAGEKEY")
    
    Set elemTemp = xmlDoc.createElement("CHARGETYPE")
    
    If intAccountNumberLength > 0 Then
        xmlSetAttributeValue elemTemp, "DATA", _
            Right(vstrAccountNumber, (intAccountNumberLength - intSeparatorPosition))
    ' RF 10/12/01
    Else
        xmlSetAttributeValue elemTemp, "DATA", "1"
    End If
    
    elemMortgagekey.appendChild elemTemp
    
    Set elemRealEstateKey = xmlDoc.createElement("REALESTATEKEY")
    Set elemRealEstateKey2 = xmlDoc.createElement("REALESTATEKEY")
    
    elemRealEstateKey2.appendChild elemRealEstateKey
    elemMortgagekey.appendChild elemRealEstateKey2
    
    
    Set elemTemp = xmlDoc.createElement("COLLATERALNUMBER")
    
    If intAccountNumberLength > 0 Then
        xmlSetAttributeValue elemTemp, "DATA", _
            Left(vstrAccountNumber, (intSeparatorPosition - 1))
    ' RF 10/12/01
    Else
        xmlSetAttributeValue elemTemp, "DATA", ""
    End If
        
    elemRealEstateKey.appendChild elemTemp
    
    Set CreateMortgageKey = elemMortgagekey
        
CreateMortgageKeyExit:

    errCheckError strFunctionName

End Function

Public Function PlexusHomeImpl_findByPrimaryKey_PrimaryCustomerKey( _
    ByVal vobjODITransformerState As ODITransformerState, _
    ByVal vnodePrimaryCustomerKey As IXMLDOMNode, _
    Optional ByVal objODIConverter As ODICONVERTER.ODICONVERTER = Nothing _
    ) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Call PlexusHomeImpl.findByPrimaryKey supplying a PrimaryCustomerKey
'   Converter request format:
'   <REQUEST>
'       <ARGUMENTS>
'           <OBJECT DATA=""1"">
'               <PRIMARYCUSTOMERKEY>
'                   <ENTITYNUMBER DATA="">
'               </PRIMARYCUSTOMERKEY>
'           </OBJECT>
'       </ARGUMENTS>
'       <METHOD DATA=""findByPrimaryKey""/>
'       <SESSION>
'           ...
'       </SESSION>
'       <TARGET DATA=""PlexusHomeImpl""/>
'   </REQUEST>"
' pass:
' return:
' exceptions:
'------------------------------------------------------------------------------------------
On Error GoTo PlexusHomeImpl_findByPrimaryKey_PrimaryCustomerKeyExit
    
    Const strFunctionName As String = _
        "PlexusHomeImpl_findByPrimaryKey_PrimaryCustomerKey"
        
    Dim nodeConverterRequest As IXMLDOMNode
    
    Set nodeConverterRequest = CreateConverterRequest( _
        vobjODITransformerState, 1, vnodePrimaryCustomerKey, _
        "findByPrimaryKey", "PlexusHomeImpl")
        
    Set PlexusHomeImpl_findByPrimaryKey_PrimaryCustomerKey = _
        SendRequestToConverter(nodeConverterRequest, objODIConverter)
        
PlexusHomeImpl_findByPrimaryKey_PrimaryCustomerKeyExit:

    Set nodeConverterRequest = Nothing

    errCheckError strFunctionName
    
End Function

Public Function PlexusHomeImpl_searchByPattern_CustomerInvolvementPattern( _
    ByVal vobjODITransformerState As ODITransformerState, _
    ByVal vstrEntityNumber As String, _
    Optional ByVal objODIConverter As ODICONVERTER.ODICONVERTER = Nothing _
    ) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   retrieves the accounts for a customer
'   Response format
'        Response
'            Argument
'            N arguments
'            CustomerInvolvementShortcut
'            Description
'            PrimaryCustomerKey
'                entityNumber
'            CISearchDescription
'                currentOutstandingBalance
'                expirySuffixDescription
'                expiryTypeDescription
'                firstAdvanceDate (e.g. = "199990801")
'                Status
'                MortgageKey
'                    ChargeType
'                    RealEstateKey
'                        CollateralNumber
'                CustomerInvolvementShortcut
'                    [etc]
' pass:
' return:
' exceptions:
'------------------------------------------------------------------------------------------
On Error GoTo PlexusHomeImpl_searchByPattern_CustomerInvolvementPatternError
    
    Const strFunctionName As String = _
        "PlexusHomeImpl_searchByPattern_CustomerInvolvementPattern"
        
    Dim nodeConverterRequest As IXMLDOMNode
        
    Set nodeConverterRequest = CreateConverterRequest( _
        vobjODITransformerState, 1, CreateCustomerInvolvementPattern(vstrEntityNumber), _
        "searchByPattern", "PlexusHomeImpl")

    Set PlexusHomeImpl_searchByPattern_CustomerInvolvementPattern = _
        SendRequestToConverter(nodeConverterRequest, objODIConverter)
    
PlexusHomeImpl_searchByPattern_CustomerInvolvementPatternExit:

    'set nodeConverterRequest = Nothing
    Exit Function

PlexusHomeImpl_searchByPattern_CustomerInvolvementPatternError:
    errCheckError strFunctionName

End Function

Public Function PlexusHomeImpl_searchByPattern_PrimaryCustomerPattern( _
    ByVal vobjODITransformerState As ODITransformerState, _
    ByVal vnodePrimaryCustomerPattern As IXMLDOMNode, _
    Optional ByVal objODIConverter As ODICONVERTER.ODICONVERTER = Nothing _
    ) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Converter request format:
'   <REQUEST PORT="" ODIENVIRONMENT="">
'       <ARGUMENTS>
'           <OBJECT DATA=""1"">
'               <PRIMARYCUSTOMERPATTERN>
'                   <FORENAME DATA=""/>
'                   <BIRTHDATE DATA=""/>
'                   <SURNAME DATA=""/>
'                   <POSTALCODE DATA=""/>
'               </PRIMARYCUSTOMERPATTERN>
'           </OBJECT>
'       </ARGUMENTS>
'       <METHOD DATA=""searchByPattern""/>
'       <SESSION>
'           ...
'       </SESSION>
'       <TARGET DATA=""PlexusHomeImpl""/>
'   </REQUEST>"
' pass:
' return:
' exceptions:
'------------------------------------------------------------------------------------------
On Error GoTo PlexusHomeImpl_searchByPattern_PrimaryCustomerPatternExit
    
    Const strFunctionName As String = _
        "PlexusHomeImpl_searchByPattern_PrimaryCustomerPattern"
            
    Dim nodeConverterRequest As IXMLDOMNode
    
    Set nodeConverterRequest = CreateConverterRequest( _
        vobjODITransformerState, 1, vnodePrimaryCustomerPattern, _
        "searchByPattern", "PlexusHomeImpl")
    
    Set PlexusHomeImpl_searchByPattern_PrimaryCustomerPattern = _
        SendRequestToConverter(nodeConverterRequest, objODIConverter)

PlexusHomeImpl_searchByPattern_PrimaryCustomerPatternExit:

    Set nodeConverterRequest = Nothing

    errCheckError strFunctionName
    
End Function

Private Function CreateConverterRequest( _
    ByVal vobjODITransformerState As ODITransformerState, _
    ByVal vintNumArguments As Integer, _
    ByVal vnodeArgument As IXMLDOMNode, _
    ByVal vstrMethodName As String, _
    ByVal vstrTargetName As String) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Build a Converter request.
'   Converter request format:
'   <REQUEST OPERATION="OPTIMUSREQUEST" OPTIMUSPORT="" ODIENVIRONMENT="">
'       <ARGUMENTS>
'           <OBJECT DATA=""1"">
'               <OPTIMUS_OBJECT>        e.g. PRIMARYCUSTOMERIMPL
'                   ...
'               </OPTIMUS_OBJECT>
'           </OBJECT>
'       </ARGUMENTS>
'       <METHOD DATA=""searchByPattern""/>
'       <SESSION>
'           ...
'       </SESSION>
'       <TARGET DATA=""PlexusHomeImpl""/>
'   </REQUEST>"
' pass:
' return:
' exceptions:
'------------------------------------------------------------------------------------------
On Error GoTo CreateConverterRequestExit

    Const strFunctionName = "CreateConverterRequest"
    
    Dim strRequest As String
    Dim xmlDoc As FreeThreadedDOMDocument
    Dim nodeRequest As IXMLDOMNode
    Dim nodeArgs  As IXMLDOMNode
    Dim elemTemp As IXMLDOMElement
    Dim nodeArgsArray As IXMLDOMNode
    Dim nodeMethod As IXMLDOMNode
    Dim nodeTarget As IXMLDOMNode
        
    Set xmlDoc = New FreeThreadedDOMDocument
    Set elemTemp = xmlDoc.createElement("REQUEST")
    Set nodeRequest = xmlDoc.appendChild(elemTemp)
    
    'AS 04/10/01 Converter request now conforms to Omiga IV Phase 2 standards.
    xmlSetAttributeValue nodeRequest, "OPERATION", "OPTIMUSREQUEST"

    'RF 17/09/01 Change to Converter Request header.
    'xmlSetAttributeValue nodeRequest, "PORT", vobjODITransformerState.GetPort()
    xmlSetAttributeValue nodeRequest, "OPTIMUSPORT", vobjODITransformerState.GetPort()
    
    xmlSetAttributeValue nodeRequest, "ODIENVIRONMENT", _
        vobjODITransformerState.GetODIEnvironment()
    
    ' arguments
    
    Set elemTemp = xmlDoc.createElement("ARGUMENTS")
    Set nodeArgs = nodeRequest.appendChild(elemTemp)
    
    Set elemTemp = xmlDoc.createElement("OBJECT")
    Set nodeArgsArray = nodeArgs.appendChild(elemTemp)
    
    Select Case vintNumArguments
    Case 0
        xmlSetAttributeValue nodeArgsArray, "DATA", "0"
    Case 1
        xmlSetAttributeValue nodeArgsArray, "DATA", "1"
        nodeArgsArray.appendChild vnodeArgument
    Case Else
        errThrowError strFunctionName, oeNotImplemented
    End Select
    
    ' method
    
    Set elemTemp = xmlDoc.createElement("METHOD")
    Set nodeMethod = nodeRequest.appendChild(elemTemp)
    xmlSetAttributeValue nodeMethod, "DATA", vstrMethodName
    
    ' session
    
    nodeRequest.appendChild vobjODITransformerState.GetSession()
    
    ' target
    
    Set elemTemp = xmlDoc.createElement("TARGET")
    Set nodeTarget = nodeRequest.appendChild(elemTemp)
    xmlSetAttributeValue nodeTarget, "DATA", vstrTargetName
        
    Set CreateConverterRequest = nodeRequest
        
CreateConverterRequestExit:

    errCheckError strFunctionName

End Function

Public Function PlexusHomeImpl_searchByPattern_CustomerListPattern( _
    ByVal vobjODITransformerState As ODITransformerState, _
    ByVal vstrAccountNumber As String, _
    Optional ByVal objODIConverter As ODICONVERTER.ODICONVERTER = Nothing _
    ) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Retrieve customers attached to a mortgage
'    Request arguement sent as a CustomerListPattern:
'        CustomerListPattern
'            MortgageKey
'                RealEstateKey
'                    CollateralNumber
'                chargeType
'   Response returned as an array of CustomerListShortcuts:
'       CustomerListShortcut
'           Description
'           PrimaryCustomerKey
'               EntityNumber
'           CLSearchDescription
'               expirySuffixDescription
'               expiryType
'               expiryTypeDescription
'               nameDetail
'                   givenLine1
'                   givenLine2
'                   givenLine3
'                   surname
'                   Title
'               PrimaryCustomerKey
'                   EntityNumber
' pass:
'   vstrAccountNumber
'       Format: collateralNumber + "." + chargeType
' return:
' exceptions:
'------------------------------------------------------------------------------------------
On Error GoTo PlexusHomeImpl_searchByPattern_CustomerListPatternExit
    
    Const strFunctionName As String = _
        "PlexusHomeImpl_searchByPattern_CustomerListPattern"
        
    Dim nodeConverterRequest As IXMLDOMNode
    Dim xmlMortgageKeyNode As IXMLDOMNode
    Dim xmlTargetElement As IXMLDOMNode
    Dim xmlCustomerList As IXMLDOMNode
    Dim xmlDoc As New FreeThreadedDOMDocument
   
    Set xmlMortgageKeyNode = CreateMortgageKey(vstrAccountNumber)
    Set xmlCustomerList = xmlDoc.createElement("CUSTOMERLISTPATTERN")
    Set xmlTargetElement = xmlDoc.createElement("TARGETKEY")
    xmlTargetElement.appendChild xmlMortgageKeyNode
    xmlCustomerList.appendChild xmlTargetElement
        
    Set nodeConverterRequest = CreateConverterRequest( _
        vobjODITransformerState, 1, xmlCustomerList, _
        "searchByPattern", "PlexusHomeImpl")
    
    Set PlexusHomeImpl_searchByPattern_CustomerListPattern = _
        SendRequestToConverter(nodeConverterRequest, objODIConverter)
    
PlexusHomeImpl_searchByPattern_CustomerListPatternExit:

    Set nodeConverterRequest = Nothing

    errCheckError strFunctionName
    
End Function

Public Function PlexusServerGateway_startNewSession( _
    ByVal vobjODITransformerState As ODITransformerState, _
    ByVal vnodeSignOnProfile As IXMLDOMNode, _
    Optional ByVal vobjODIConverter As ODICONVERTER.ODICONVERTER = Nothing _
    ) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Call PlexusServerGateway.startNewSession to return a new Session
'   Converter request format:
'   <REQUEST>
'       <ARGUMENTS>
'           <OBJECT DATA=""1"">
'               <SIGNONPROFILE>
'                   <HOST DATA=""/>
'                   <ENVIRONMENT DATA=""/>
'                   <USERNAME DATA=""/>
'                   <PASSWORD DATA=""/>
'               </SIGNONPROFILE>
'           </OBJECT>
'       </ARGUMENTS>
'       <METHOD DATA=""startNewSession""/>
'       <SESSION>
'           ...
'       </SESSION>
'       <TARGET DATA=""PlexusServerGatewayImpl""/>
'   </REQUEST>"
' pass:
' return:
' exceptions:
'------------------------------------------------------------------------------------------
On Error GoTo PlexusServerGateway_startNewSessionExit

    Const strFunctionName = "PlexusServerGateway_startNewSession"
    
    Dim nodeConverterRequest As IXMLDOMNode
    
    Set nodeConverterRequest = CreateConverterRequest( _
        vobjODITransformerState, 1, vnodeSignOnProfile, _
        "startNewSession", "PlexusServerGatewayImpl")
        
    Set PlexusServerGateway_startNewSession = _
        SendRequestToConverter(nodeConverterRequest, vobjODIConverter)
    
PlexusServerGateway_startNewSessionExit:

    Set nodeConverterRequest = Nothing

    errCheckError strFunctionName
    
End Function

Public Function Session_endSession( _
    ByVal vobjODITransformerState As ODITransformerState, _
    Optional ByVal vobjODIConverter As ODICONVERTER.ODICONVERTER = Nothing _
    ) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Call Session.endSession to end the Session
'   Converter request format:
'   <REQUEST>
'       <ARGUMENTS>
'           <OBJECT DATA=""0""/>
'       </ARGUMENTS>
'       <METHOD DATA=""endSession""/>
'       <SESSION>
'           ...
'       </SESSION>
'       <TARGET DATA=""SessionImpl""/>
'   </REQUEST>"
' pass:
' return:
' exceptions:
'------------------------------------------------------------------------------------------
On Error GoTo Session_endSessionExit
    
    Const strFunctionName As String = "Session_endSession"
        
    Dim nodeConverterRequest As IXMLDOMNode
    Dim nodeDummy As IXMLDOMNode
    
    Set nodeConverterRequest = CreateConverterRequest( _
        vobjODITransformerState, 0, nodeDummy, "endSession", "SessionImpl")

    Set Session_endSession = _
        SendRequestToConverter(nodeConverterRequest, vobjODIConverter)
    
Session_endSessionExit:

    Set nodeConverterRequest = Nothing

    errCheckError strFunctionName

End Function

Public Function PlexusHomeImpl_findByPrimaryKey_MortgageKey( _
    ByVal vobjODITransformerState As ODITransformerState, _
    ByVal vnodeMortgageKey As IXMLDOMNode, _
    Optional ByVal objODIConverter As ODICONVERTER.ODICONVERTER = Nothing _
    ) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Call PlexusHomeImpl.findByPrimaryKey supplying a MortgageKey
' pass:
' return:
' exceptions:
'------------------------------------------------------------------------------------------
On Error GoTo PlexusHomeImpl_findByPrimaryKey_MortgageKeyExit
    
    Const strFunctionName As String = _
        "PlexusHomeImpl_findByPrimaryKey_MortgageKey"
        
    Dim nodeConverterRequest As IXMLDOMNode
    
    Set nodeConverterRequest = CreateConverterRequest( _
        vobjODITransformerState, 1, vnodeMortgageKey, _
        "findByPrimaryKey", "PlexusHomeImpl")
        
    Set PlexusHomeImpl_findByPrimaryKey_MortgageKey = _
        SendRequestToConverter(nodeConverterRequest, objODIConverter)
        
PlexusHomeImpl_findByPrimaryKey_MortgageKeyExit:

    Set nodeConverterRequest = Nothing

    errCheckError strFunctionName
    
End Function

Public Function TableSnifferImpl_getTable( _
    ByVal vobjODITransformerState As ODITransformerState, _
    ByVal vstrTableName As IXMLDOMNode, _
    Optional ByVal objODIConverter As ODICONVERTER.ODICONVERTER = Nothing _
    ) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Call TableSnifferImpl.getTable
' pass:
' return:
' exceptions:
'------------------------------------------------------------------------------------------
On Error GoTo TableSnifferImpl_getTableExit
    
    Const strFunctionName As String = _
        "TableSnifferImpl_getTable"
        
    Dim nodeConverterRequest As IXMLDOMNode
    
    ' fixme - check whether argument is a string or really is a TableImpl
    Set nodeConverterRequest = CreateConverterRequest( _
        vobjODITransformerState, 1, vstrTableName, _
        "getTable", "TableSnifferImpl")
        
    Set TableSnifferImpl_getTable = _
        SendRequestToConverter(nodeConverterRequest, objODIConverter)
        
TableSnifferImpl_getTableExit:

    Set nodeConverterRequest = Nothing

    errCheckError strFunctionName
    
End Function
'DM     03/04/02    SYS4350 Implement SaveThirdpartyDetails
Public Function PrimaryCustomerImpl_update( _
    ByVal vobjODITransformerState As ODITransformerState, _
    ByVal vnodePrimaryCustomer As IXMLDOMNode, _
    Optional ByVal objODIConverter As ODICONVERTER.ODICONVERTER = Nothing _
    ) As IXMLDOMNode

    Const strFunctionName As String = "PrimaryCustomerImpl_update"
    
    Dim nodeConverterRequest As IXMLDOMNode
    
    ' this needs some work yet.
    Set nodeConverterRequest = CreateConverterRequest( _
        vobjODITransformerState, 1, vnodePrimaryCustomer, "update", "PrimaryCustomerImpl")
    
    ' Send the data to the converter.
    Set PrimaryCustomerImpl_update = SendRequestToConverter(nodeConverterRequest, objODIConverter)
    
PrimaryCustomerImpl_update:

    Set nodeConverterRequest = Nothing
    
    errCheckError strFunctionName

End Function


'SYS4609 - Create a Mortgage Product in Optimus.
Public Function PlexusHomeImpl_create_MortgageProductKey( _
    ByVal vobjODITransformerState As ODITransformerState, _
    ByVal sMortgageProductCode As String) As IXMLDOMNode

    Dim xmlConverterRequest As IXMLDOMNode

    Const strFunctionName As String = "PlexusHomeImpl_create_MortgageProductKey"
    
    On Error GoTo PlexusHomeImpl_create_MortgageProductKeyExit
    
    'Convert the request.
    Set xmlConverterRequest = CreateConverterRequest( _
        vobjODITransformerState, 1, CreateMortgageProductKey(sMortgageProductCode), "create", "PrimaryCustomerImpl")
    
    'Send the request to the converter.
    Set PlexusHomeImpl_create_MortgageProductKey = SendRequestToConverter(xmlConverterRequest)
    
PlexusHomeImpl_create_MortgageProductKeyExit:
    Set xmlConverterRequest = Nothing

    errCheckError strFunctionName
End Function


'SYS4609 - Return a <MORTGAGEPRODUCTKEY DATA=""/> element.
Private Function CreateMortgageProductKey(ByVal sMortgageProductCode As String) As IXMLDOMNode

    Dim xmlElement As IXMLDOMNode
    Dim xmlMortgageProductKey As IXMLDOMNode
    Dim xmlDocument As FreeThreadedDOMDocument
    
    Set xmlDocument = New FreeThreadedDOMDocument
    xmlDocument.async = False
    
    Set xmlMortgageProductKey = xmlDocument.createElement("MORTGAGEPRODUCTKEY")
    Set xmlElement = xmlDocument.createElement("ENTITYNUMBER")
    xmlMortgageProductKey.appendChild xmlElement
    
    xmlSetAttributeValue xmlElement, "DATA", sMortgageProductCode
    
    Set CreateMortgageProductKey = xmlMortgageProductKey

End Function


'SYS4609 - Update the Mortgage Product in Optimus.
Public Function MortgageProductImpl_update( _
    ByVal vobjODITransformerState As ODITransformerState, _
    ByVal xmlMortgageProductImpl As IXMLDOMNode) As IXMLDOMNode

    Dim xmlConverterRequest As IXMLDOMNode

    Const strFunctionName As String = "MortgageProductImpl_update"
    
    On Error GoTo MortgageProductImpl_updateExit
    
    'Convert the request.
    Set xmlConverterRequest = CreateConverterRequest( _
        vobjODITransformerState, 1, xmlMortgageProductImpl, "update", "MortgageProductImpl")
    
    'Send the request to the converter.
    Set MortgageProductImpl_update = SendRequestToConverter(xmlConverterRequest)
    
MortgageProductImpl_updateExit:
    Set xmlConverterRequest = Nothing
    
    errCheckError strFunctionName
End Function


'SYS4609 - Create a User Key in Optimus.
Public Function PlexusHomeImpl_create_UserProfileKey( _
    ByVal vobjODITransformerState As ODITransformerState, _
    ByVal sUserID As String) As IXMLDOMNode

    Dim xmlConverterRequest As IXMLDOMNode

    Const strFunctionName As String = "PlexusHomeImpl_create_UserProfileKey"
    
    On Error GoTo PlexusHomeImpl_create_UserProfileKeyExit
    
    'Convert the request.
    Set xmlConverterRequest = CreateConverterRequest( _
        vobjODITransformerState, 1, CreateUserProfileKey(sUserID), "create", "UserProfileImpl")
    
    'Send the request to the converter.
    Set PlexusHomeImpl_create_UserProfileKey = SendRequestToConverter(xmlConverterRequest)
    
PlexusHomeImpl_create_UserProfileKeyExit:
    Set xmlConverterRequest = Nothing

    errCheckError strFunctionName
End Function


'SYS4609 - Return a <USERPROFILEKEY DATA=""/> element.
Public Function CreateUserProfileKey(ByVal sUserID As String) As IXMLDOMNode

    Dim xmlElement As IXMLDOMNode
    Dim xmlUserProfileKey As IXMLDOMNode
    Dim xmlDocument As FreeThreadedDOMDocument

    Set xmlDocument = New FreeThreadedDOMDocument
    xmlDocument.async = False

    Set xmlUserProfileKey = xmlDocument.createElement("USERPROFILEKEY")
    Set xmlElement = xmlDocument.createElement("ENTITYNUMBER")
    xmlUserProfileKey.appendChild xmlElement
    xmlSetAttributeValue xmlElement, "DATA", sUserID

    Set CreateUserProfileKey = xmlUserProfileKey

End Function


'SYS4609 - Update the user in Optimus.
Public Function UserProfileImpl_update( _
    ByVal vobjODITransformerState As ODITransformerState, _
    ByVal xmlUserProfileImpl As IXMLDOMNode) As IXMLDOMNode
    Dim xmlConverterRequest As IXMLDOMNode

    Const strFunctionName As String = "UserProfileImpl_update"
    
    On Error GoTo UserProfileImpl_updateExit
    
    'Convert the request.
    Set xmlConverterRequest = CreateConverterRequest( _
        vobjODITransformerState, 1, xmlUserProfileImpl, "update", "UserProfileImpl")
    
    'Send the request to the converter.
    Set UserProfileImpl_update = SendRequestToConverter(xmlConverterRequest)
    
UserProfileImpl_updateExit:
    Set xmlConverterRequest = Nothing
    
    errCheckError strFunctionName
    
End Function


'SYS4609 - Find a user in Optimus.
Public Function PlexusHomeImpl_findByPrimaryKey_UserProfileKey( _
    ByVal vobjODITransformerState As ODITransformerState, _
    ByVal vnodeUserProfileKey As IXMLDOMNode, _
    Optional ByVal objODIConverter As ODICONVERTER.ODICONVERTER = Nothing _
    ) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Call PlexusHomeImpl.findByPrimaryKey supplying a UserProfileKey
'   Converter request format:
'   <REQUEST>
'       <ARGUMENTS>
'           <OBJECT DATA=""1"">
'               <USERPROFILEKEY DATA="">
'           </OBJECT>
'       </ARGUMENTS>
'       <METHOD DATA=""findByPrimaryKey""/>
'       <SESSION>
'           ...
'       </SESSION>
'       <TARGET DATA=""PlexusHomeImpl""/>
'   </REQUEST>"
' pass:
' return:
' exceptions:
'------------------------------------------------------------------------------------------
On Error GoTo PlexusHomeImpl_findByPrimaryKey_UserProfileKeyExit
    
    Const strFunctionName As String = "PlexusHomeImpl_findByPrimaryKey_UserProfileKey"
        
    Dim nodeConverterRequest As IXMLDOMNode
    
    Set nodeConverterRequest = CreateConverterRequest( _
        vobjODITransformerState, 1, vnodeUserProfileKey, _
        "findByPrimaryKey", "PlexusHomeImpl")
        
    Set PlexusHomeImpl_findByPrimaryKey_UserProfileKey = _
        SendRequestToConverter(nodeConverterRequest, objODIConverter)
        
PlexusHomeImpl_findByPrimaryKey_UserProfileKeyExit:

    Set nodeConverterRequest = Nothing

    errCheckError strFunctionName
    
End Function


'SYS5101 - Send the rate information to Optimus.
Public Function PlexusHomeImpl_ProcessRateChange( _
    ByVal vobjODITransformerState As ODITransformerState, _
    ByVal vnodeRateChangeCall As IXMLDOMNode, _
    Optional ByVal objODIConverter As ODICONVERTER.ODICONVERTER = Nothing) As IXMLDOMNode

    Const strFunctionName As String = "PlexusHomeImpl_ProcessRateChange"

    Dim nodeConverterRequest As IXMLDOMNode

    Set nodeConverterRequest = CreateConverterRequest(vobjODITransformerState, 1, vnodeRateChangeCall, "processRateChange", "PlexusHomeImpl")
        
    Set PlexusHomeImpl_ProcessRateChange = SendRequestToConverter(nodeConverterRequest, objODIConverter)
        
PelxusHomeImpl_ProcessRateChangeExit:
    Set nodeConverterRequest = Nothing

    errCheckError strFunctionName

End Function
    
'SG/JR 21/11/02 SYS5765/SYSMCP1256
Public Function PlexusHomeImpl_findByPrimaryKey_ThirdPartyKey(ByVal vobjODITransformerState As ODITransformerState, _
    ByVal vnodeThirdPartyKey As IXMLDOMNode, _
    Optional ByVal objODIConverter As ODICONVERTER.ODICONVERTER = Nothing _
    ) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'   Call PlexusHomeImpl.findByPrimaryKey supplying a ThirdPartyKey
' pass:
' return:
' exceptions:
'------------------------------------------------------------------------------------------
On Error GoTo PlexusHomeImpl_findByPrimaryKey_ThirdPartyKeyExit
    
    Const strFunctionName As String = _
        "PlexusHomeImpl_findByPrimaryKey_ThirdPartyKey"
        
    Dim nodeConverterRequest As IXMLDOMNode
    
    Set nodeConverterRequest = CreateConverterRequest( _
        vobjODITransformerState, 1, vnodeThirdPartyKey, _
        "findByPrimaryKey", "PlexusHomeImpl")
        
    Set PlexusHomeImpl_findByPrimaryKey_ThirdPartyKey = _
        SendRequestToConverter(nodeConverterRequest, objODIConverter)
        
PlexusHomeImpl_findByPrimaryKey_ThirdPartyKeyExit:

    Set nodeConverterRequest = Nothing

    errCheckError strFunctionName

End Function
'End



