Attribute VB_Name = "omTMFunctions"
'Workfile:      CommonFunctions.bas
'Copyright:     Copyright © 2005 Marlborough Stirling
'Description:   Functions common to other modules.
'
'-------------------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'GHun   01/08/2005  MAR7 Apply CORE174
'PSC    26/08/2005  MAR32 Add GetApplicationOwners
'PSC    24/01/2006  MAR1118 Add IsDebugEnabled
'PSC    28/02/2006  MAR1341 Add Chasing Period Minutes
'HMA    06/04/2006  MAR1408 Add optional parameter to SetCaseTaskDueDateAndTime
'LH     05/01/2007  EP2_352 Remote ownership changes
'IK     15/01/2007  EP2_861 fix to IsCaseOwnerRemote for CreateAdhocCaseTask
'IK     20/01/2007  EP2_902 AddDefaultValuesToCaseStage moved to omTmFunctions.bas
'IK     10/04/2007  EP2_2227 fix to AddDefaultValuesToCaseTask for CreateAdhocCaseTask
'------------------------------------------------------------------------------------------------------

Option Explicit

Public Sub AddDefaultValuesToCaseStage( _
    ByRef objContext As ObjectContext, _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlNextCaseStageNode As IXMLDOMNode, _
    ByVal vxmlNextStageNode As IXMLDOMNode)
    
    On Error GoTo AddDefaultValuesToCaseStageExit
    Const cstrFunctionName As String = "AddDefaultValuesToCaseStage"
    
    Dim xmlCaseTaskNode As IXMLDOMNode
    Dim xmlTaskNode As IXMLDOMNode
    
    Dim strPattern As String
    
    Dim strUserId As String
    Dim strUnitId As String
    Call GetApplicationOwners(objContext, vxmlRequestNode, strUserId, strUnitId)
    
    'EP2_902 big frig to save repeated calls to IsCaseOwnerRemote
    Dim objOwnerDetails As Collection
    Dim blnIncludeRemoteOwnership As Boolean
    Dim blnRemoteUnderwriter As Boolean

    blnIncludeRemoteOwnership = GetGlobalParamBoolean("TMIncludeRemoteOwnership")
    
    If blnIncludeRemoteOwnership = True Then
        blnRemoteUnderwriter = _
            IsCaseOwnerRemote(vxmlRequestNode, strUnitId, strUserId)
    End If
    
    Set objOwnerDetails = New Collection
    objOwnerDetails.Add blnIncludeRemoteOwnership, "isIncludeRemoteOwnership"
    objOwnerDetails.Add blnRemoteUnderwriter, "isRemoteUnderwriter"
    
    For Each xmlCaseTaskNode In vxmlNextCaseStageNode.childNodes
    
        strPattern = _
            "STAGETASK[@TASKID='" & xmlGetAttributeText(xmlCaseTaskNode, "TASKID") & "']"
        
        Set xmlTaskNode = vxmlNextStageNode.selectSingleNode(strPattern)
        
        'MAR7 GHun AddDefaultValuesToCaseTask moved to omTMFunctions.bas; pass gobjContext.
        AddDefaultValuesToCaseTask objContext, vxmlRequestNode, xmlCaseTaskNode, xmlTaskNode, strUserId, strUnitId, objOwnerDetails
        
        ' DRC AQR SYS2266 This Call Moved
        '  ProcessAutomaticTasks vxmlRequestNode, xmlCaseTaskNode
               
        ' AQR SYS1791
        ' add ORIGINATINGSTAGEID
        xmlCopyAttributeValue xmlTaskNode, xmlCaseTaskNode, "STAGEID", "ORIGINATINGSTAGEID"
                                
    Next
    
AddDefaultValuesToCaseStageExit:
    
    Set objOwnerDetails = Nothing
    Set xmlCaseTaskNode = Nothing
    Set xmlTaskNode = Nothing
    
    errCheckError cstrFunctionName

End Sub

Public Sub AddDefaultValuesToCaseTask( _
    ByRef objContext As ObjectContext, _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlCaseTaskNode As IXMLDOMNode, _
    ByVal vxmlTaskNode As IXMLDOMNode, _
    ByVal vstrDefaultUserId As String, _
    ByVal vStrDefaultUnitId As String, _
    Optional ByVal vobjobjOwnerDetails As Collection = Nothing)
    
On Error GoTo AddDefaultValuesToCaseTaskExit
    
    Const cstrFunctionName As String = "AddDefaultValuesToCaseTask"
    
    Dim strTaskUserID       As String
    Dim strUnitId           As String
    Dim strTaskOwnerType    As String
    Dim strTmpResponse      As String
    Dim strOwningUserID     As String
    Dim strOwningUnitID     As String
    Dim strCaseTaskUserID   As String
    Dim strCaseTaskUnitID   As String
    
    Dim objOrg              As Object
      
    Dim xmlIn               As FreeThreadedDOMDocument40
    Dim xmlOut              As FreeThreadedDOMDocument40
    Dim xmlElement          As IXMLDOMElement
    Dim xmlTempNode         As IXMLDOMNode
    Dim xmlParentNode       As IXMLDOMNode
    Dim xmlChildNode        As IXMLDOMNode
    Dim lngErrNo            As Long
    
    Dim blRemoteOwnerTaskInd As Boolean
    Dim blnRemoteUnderwriter As Boolean
    Dim blnIncludeRemoteOwnership As Boolean
    
    Set xmlIn = New FreeThreadedDOMDocument40
    xmlIn.validateOnParse = False
    xmlIn.setProperty "NewParser", True
    Set xmlOut = New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    
    
    Set xmlElement = vxmlRequestNode
    Set xmlParentNode = xmlIn.appendChild(xmlElement)
    If xmlAttributeValueExists(vxmlCaseTaskNode, "CASETASKNAME") = False Then
        xmlCopyAttributeValue vxmlTaskNode, vxmlCaseTaskNode, "TASKNAME", "CASETASKNAME"
    End If
            
    If xmlAttributeValueExists(vxmlCaseTaskNode, "TASKDUEDATEANDTIME") = False Then
        SetCaseTaskDueDateAndTime vxmlCaseTaskNode, vxmlTaskNode
    End If
    
    If xmlGetAttributeText(vxmlTaskNode, "REMOTEOWNERTASKIND") = "1" Then
        blRemoteOwnerTaskInd = True
        If vobjobjOwnerDetails Is Nothing Then
            blnIncludeRemoteOwnership = GetGlobalParamBoolean("TMIncludeRemoteOwnership")
        Else
            blnIncludeRemoteOwnership = vobjobjOwnerDetails.Item("isIncludeRemoteOwnership")
        End If
        If blnIncludeRemoteOwnership Then
            If vobjobjOwnerDetails Is Nothing Then
                blnRemoteUnderwriter = IsCaseOwnerRemote(vxmlRequestNode, strCaseTaskUnitID, strCaseTaskUserID)
            Else
                blnRemoteUnderwriter = vobjobjOwnerDetails.Item("isRemoteUnderwriter")
                ' EP2_2227
                If blnRemoteUnderwriter Then
                    strCaseTaskUserID = vstrDefaultUserId
                    strCaseTaskUnitID = vStrDefaultUnitId
                End If
                ' EP2_2227 ends
            End If
        End If
    End If
        
    'Assign task ownership to remote owner if applicable
    If blnIncludeRemoteOwnership = True _
    And blRemoteOwnerTaskInd = True _
    And blnRemoteUnderwriter = True _
    Then
                
        ' EP2_2227
        strOwningUserID = strCaseTaskUserID
        strOwningUnitID = strCaseTaskUnitID
        ' EP2_2227 ends
    
    Else
    
        '  Getting the OwningUnitID and OwningUserID
        strTaskUserID = xmlGetAttributeText(vxmlTaskNode, "TASKUSERID")
        
        strCaseTaskUserID = xmlGetAttributeText(vxmlCaseTaskNode, "OWNINGUSERID")
        strCaseTaskUnitID = xmlGetAttributeText(vxmlCaseTaskNode, "OWNINGUNITID")
        
        If Len(strCaseTaskUserID) > 0 Then
            'AS 25/05/2005 CORE174 Use OWNINGUSERID and OWNINGUNITID in request.
            'This enables the user to create an ad hoc task for a nominated user.
            strOwningUserID = strCaseTaskUserID
            strOwningUnitID = strCaseTaskUnitID
            
        ElseIf Len(strTaskUserID) > 0 Then
            strOwningUserID = xmlGetAttributeText(vxmlTaskNode, "TASKUSERID")
            strOwningUnitID = xmlGetAttributeText(vxmlTaskNode, "TASKUNITID")
            
        Else 'Non Stage Task User
            strTaskOwnerType = xmlGetAttributeText(vxmlTaskNode, "TASKOWNERTYPE")
            If Len(strTaskOwnerType) > 0 Then
                
                'Build up request for call to omOrg.OrganisationBO.GetActionOwnerDetails
                
                strUnitId = xmlGetAttributeText(vxmlTaskNode, "TASKUNITID")
                If Len(strUnitId) = 0 Then 'Stage Task Unit undefined
                    strUnitId = vStrDefaultUnitId
                End If
                strTaskOwnerType = xmlGetAttributeText(vxmlTaskNode, "TASKOWNERTYPE")
                'GD Testing .. strTaskOwnerType = "40"
                Set xmlChildNode = xmlIn.createElement("ACTIONOWNER")
                xmlParentNode.appendChild xmlChildNode
                Set xmlTempNode = xmlIn.createElement("UNITID")
                xmlTempNode.Text = strUnitId
                xmlChildNode.appendChild xmlTempNode
                Set xmlTempNode = xmlIn.createElement("TASKOWNERTYPE")
                xmlTempNode.Text = strTaskOwnerType
                xmlChildNode.appendChild xmlTempNode
                
                'Extract Request Info and append actionowner stuff to it..
                Set xmlTempNode = vxmlRequestNode.cloneNode(False)
                xmlTempNode.appendChild xmlChildNode
                
                ' Calling OrganisationBO.GetActionOwnerDetails
                Set objOrg = objContext.CreateInstance("omOrg.OrganisationBO")
                strTmpResponse = objOrg.GetActionOwnerDetails(xmlTempNode.xml)
                Set xmlOut = xmlLoad(strTmpResponse, cstrFunctionName)
                
                ' Error Check
                Set xmlTempNode = xmlOut.selectSingleNode("RESPONSE/ERROR")
                If Not xmlTempNode Is Nothing Then
                    lngErrNo = xmlGetNodeAsLong(xmlTempNode, "NUMBER")
                    
                    ' Check for record not found error
                    If errGetOmigaErrorNumber(lngErrNo) = oeRecordNotFound Then 'No Task Owner for Type
                        strOwningUserID = vstrDefaultUserId
                        strOwningUnitID = vStrDefaultUnitId
                    Else
                        
                        'Error other than Record Not Found occurred
                        errCheckXMLResponseNode xmlOut.documentElement, , True
                    End If
                Else
                    'No error has occurred
                    Set xmlTempNode = xmlOut.selectSingleNode("RESPONSE/ACTIONOWNER")
                    If Not xmlTempNode Is Nothing Then
                        Set xmlElement = xmlOut.createElement("ACTIONOWNER")
                        strOwningUserID = xmlGetNodeText(xmlTempNode, "USERID") 'Task Owner User
                        strOwningUnitID = xmlGetNodeText(xmlTempNode, "UNITID") 'Task Owner Unit
                    End If
                 End If
            Else 'No Task Owner
                strOwningUserID = vstrDefaultUserId
                strOwningUnitID = vStrDefaultUnitId
            End If
        End If
    End If
    xmlSetAttributeValue vxmlCaseTaskNode, "OWNINGUSERID", strOwningUserID
    xmlSetAttributeValue vxmlCaseTaskNode, "OWNINGUNITID", strOwningUnitID

    If Not xmlAttributeValueExists(vxmlCaseTaskNode, "MANDATORYINDICATOR") Then
        xmlCopyAttributeValue vxmlTaskNode, vxmlCaseTaskNode, "MANDATORYFLAG", "MANDATORYINDICATOR"
    End If
    
AddDefaultValuesToCaseTaskExit:
    
    Set objOrg = Nothing
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlTempNode = Nothing
    Set xmlElement = Nothing
    Set xmlParentNode = Nothing
    Set xmlChildNode = Nothing
    errCheckError cstrFunctionName
End Sub

'MAR1408 Add Completion Date as an optional parameter
Public Sub SetCaseTaskDueDateAndTime( _
    ByVal vxmlCaseTaskNode As IXMLDOMNode, _
    ByVal vxmlTaskNode As IXMLDOMNode, _
    Optional ByVal vstrCompletionDate As String = "")
    Const cstrFunctionName As String = "SetCaseTaskDueDateAndTime"
    
On Error GoTo SetCaseTaskDueDateAndTimeExit

    Dim dtDueDateTime       As Date
    Dim intChasePeriodDays  As Integer
    Dim intChasePeriodHours As Integer
    Dim intOffsetDays       As Integer
    Dim intOffsetHours      As Integer
    
    ' PSC 28/02/2006 MAR1341 - Start
    Dim intChasePeriodMinutes  As Integer
    Dim intOffsetMinutes       As Integer
    ' PSC 28/02/2006 MAR1341 - End
    
    'MAR1408 Set Due Date = Now unless a date has been passed in.
    If (vstrCompletionDate <> "") Then
        dtDueDateTime = CSafeDate(vstrCompletionDate)
    Else
        dtDueDateTime = Now()
    End If
    
    intChasePeriodDays = xmlGetAttributeAsInteger(vxmlTaskNode, "CHASINGPERIODDAYS")
    intChasePeriodHours = xmlGetAttributeAsInteger(vxmlTaskNode, "CHASINGPERIODHOURS")
    
    ' PSC 28/02/2006 MAR1341
    intChasePeriodMinutes = xmlGetAttributeAsInteger(vxmlTaskNode, "CHASINGPERIODMINUTES")
    
    If intChasePeriodDays > 0 Then
        dtDueDateTime = DateAdd("y", intChasePeriodDays, dtDueDateTime)
    End If
    
    If intChasePeriodHours > 0 Then
        dtDueDateTime = DateAdd("h", intChasePeriodHours, dtDueDateTime)
    End If
    
    ' PSC 28/02/2006 MAR1341 - Start
    If intChasePeriodMinutes > 0 Then
        dtDueDateTime = DateAdd("n", intChasePeriodMinutes, dtDueDateTime)
    End If
    ' PSC 28/02/2006 MAR1341 - End
    
    intOffsetDays = xmlGetAttributeAsInteger(vxmlTaskNode, "ADJUSTMENTDAYS")
    intOffsetHours = xmlGetAttributeAsInteger(vxmlTaskNode, "ADJUSTMENTHOURS")
    
    ' PSC 28/02/2006 MAR1341
    intOffsetMinutes = xmlGetAttributeAsInteger(vxmlTaskNode, "ADJUSTMENTMINUTES")
    
    If intOffsetDays <> 0 Then
        dtDueDateTime = DateAdd("y", intOffsetDays, dtDueDateTime)
    End If
    
    If intOffsetHours <> 0 Then
        dtDueDateTime = DateAdd("h", intOffsetHours, dtDueDateTime)
    End If
    
    ' PSC 28/02/2006 MAR1341 - Start
    If intOffsetMinutes <> 0 Then
        dtDueDateTime = DateAdd("n", intOffsetMinutes, dtDueDateTime)
    End If
    ' PSC 28/02/2006 MAR1341 - End
    
    If dtDueDateTime < Now() Then
        dtDueDateTime = Now()
    End If
    
    xmlSetAttributeValue vxmlCaseTaskNode, "TASKDUEDATEANDTIME", dtDueDateTime
SetCaseTaskDueDateAndTimeExit:

    errCheckError cstrFunctionName
End Sub
'PSC 25/08/2005 MAR32 - Start
Public Sub GetApplicationOwners(ByRef objContext As ObjectContext, _
                                 ByVal vxmlRequestNode As IXMLDOMNode, _
                                 ByRef strUserId As String, ByRef strUnitId As String)

On Error GoTo GetApplicationOwnersExit

Dim xmlFindAppOwnerShipListDoc As FreeThreadedDOMDocument40
Dim xmlNode As IXMLDOMNode
Dim xmlAppOwnerShipNode As IXMLDOMElement
Dim xmlAppNumberNode As IXMLDOMElement
Dim objAppManBO As Object
Dim xmlTempNodeList As IXMLDOMNode
Dim xmlTempNode As IXMLDOMNode

Dim strFindAppOwnerShipListDoc As String
Dim lngErrNo As Long
Dim dteTempDate As Date
Dim dteMaxDate As Date
Dim intListCount  As Integer
Dim intMaxIndex  As Integer
Dim intListIndex As Integer

Const cstrFunctionName As String = "GetApplicationOwners"

    Set xmlFindAppOwnerShipListDoc = New FreeThreadedDOMDocument40
    xmlFindAppOwnerShipListDoc.validateOnParse = False
    xmlFindAppOwnerShipListDoc.setProperty "NewParser", True
    xmlFindAppOwnerShipListDoc.async = False
    
    'Formulate the Request
    Set xmlNode = xmlFindAppOwnerShipListDoc.appendChild(vxmlRequestNode.cloneNode(False))
    Set xmlAppOwnerShipNode = xmlFindAppOwnerShipListDoc.createElement("APPLICATIONOWNERSHIP")
    Set xmlAppNumberNode = xmlFindAppOwnerShipListDoc.createElement("APPLICATIONNUMBER")
    
    'Search for CASEID attrib in Request
    xmlAppNumberNode.Text = xmlGetMandatoryNodeText(vxmlRequestNode, ".//*/@CASEID")
    xmlAppOwnerShipNode.appendChild xmlAppNumberNode
    xmlNode.appendChild xmlAppOwnerShipNode
    
    'Call ApplicationManagerBO
    Set objAppManBO = objContext.CreateInstance(gstrAPPLICATION_COMPONENT & ".ApplicationManagerBO")
    strFindAppOwnerShipListDoc = objAppManBO.FindApplicationOwnershipList(xmlNode.xml)
    xmlFindAppOwnerShipListDoc.loadXML strFindAppOwnerShipListDoc
    Set xmlAppOwnerShipNode = xmlFindAppOwnerShipListDoc.documentElement
    
    'Process response
    lngErrNo = errCheckXMLResponseNode(xmlAppOwnerShipNode, , False)
    If lngErrNo <> 0 Then
        'Check for record not found error
        If errGetOmigaErrorNumber(lngErrNo) = oeRecordNotFound Then
            'No Application owner
            strUserId = xmlGetMandatoryAttributeText(vxmlRequestNode, "USERID")
            strUnitId = xmlGetMandatoryAttributeText(vxmlRequestNode, "UNITID")
        Else
            'raise error and exit
            errCheckXMLResponseNode xmlAppOwnerShipNode, , True
        End If
    Else
        'Find the latest owner details
        Set xmlTempNodeList = xmlGetMandatoryNode(xmlAppOwnerShipNode, ".//USERHISTORYLIST")
        dteMaxDate = "01/01/1800"
        intListCount = xmlTempNodeList.childNodes.length
        
        If intListCount > 0 Then
            For intListIndex = 0 To (intListCount - 1)
                'Set xmlTempNode = xmlTempNodeList.childNodes.Item(intListIndex).selectSingleNode("USERHISTORYDATE")
                'dteTempDate = CDate(xmlTempNode.Text)
                dteTempDate = xmlGetMandatoryNodeAsDate(xmlTempNodeList.childNodes.Item(intListIndex), "USERHISTORYDATE")
                If dteTempDate > dteMaxDate Then
                    dteMaxDate = dteTempDate
                    intMaxIndex = intListIndex ' The item number of the USERHISTORY which we want
                End If
            Next
            'Extract User and Unit ID
            strUserId = xmlGetNodeText(xmlTempNodeList.childNodes.Item(intMaxIndex), ".//USERID")
            strUnitId = xmlGetNodeText(xmlTempNodeList.childNodes.Item(intMaxIndex), ".//UNITID")
        Else
            'No Application owner
            strUserId = xmlGetMandatoryAttributeText(vxmlRequestNode, "USERID")
            strUnitId = xmlGetMandatoryAttributeText(vxmlRequestNode, "UNITID")
        End If
        
    End If
        
GetApplicationOwnersExit:
    Set xmlFindAppOwnerShipListDoc = Nothing
    Set xmlNode = Nothing
    Set xmlAppOwnerShipNode = Nothing
    Set xmlAppNumberNode = Nothing
    Set objAppManBO = Nothing
    Set xmlTempNodeList = Nothing
    Set xmlTempNode = Nothing
    
    errCheckError cstrFunctionName

End Sub
'PSC 25/08/2005 MAR32 - End

'PSC 24/01/2006 MAR1118 - Start
Public Function IsDebugEnabled() As Boolean
    
    Const strRegSection As String = "HKLM\SOFTWARE\Omiga4\System Configuration\"
    
    Dim objWshShell     As Object
    Dim strDebugEnabled    As String
       
    ' Ignore any errors that occur
    On Error Resume Next
    
    Set objWshShell = CreateObject("WScript.Shell")
    
    ' Read the path to save admin debugging info to from the registry
    strDebugEnabled = Trim(objWshShell.RegRead(strRegSection & "omTMDebug"))
    
    If Len(strDebugEnabled) > 0 Then
        If UCase(strDebugEnabled) = "Y" Or strDebugEnabled = "1" Then
            IsDebugEnabled = True
        End If
    End If
    
    'Clear any errors that may have occurred as they can be ignored
    Err.Clear
End Function
'PSC 24/01/2006 MAR1118 - End

'LH 05/01/2007
'Checks if the case owner is a remote owner (underwriter) within their unit.
'Returns true if the case owner is a remote owner, otherwise false.
'If the case owner is a remote owner, the UnitID & UserID are assigned & returned
'through their corresponding parameters (rstrUnitID & rstrUserID)
Private Function IsCaseOwnerRemote(ByVal vxmlRequest As IXMLDOMNode, _
                                    ByRef rstrUnitID As String, _
                                    ByRef rstrUserID As String) As Boolean

    Const cstrFunctionName As String = "IsCaseOwnerRemote"
    On Error GoTo IsCaseOwnerRemote_Error
    
    Dim xmlDoc As DOMDocument40
    Dim xmlElem As IXMLDOMElement
    Dim xmlNode As IXMLDOMElement
    Dim xmlApplicationNode As IXMLDOMNode
    Dim xmlReturnDoc As FreeThreadedDOMDocument40
    Dim objCRUD As Object
    Dim blSuccess As Boolean
    
    Dim strApplicationNumber As String
    Dim strApplicationFFNumber As String
    Dim strWithinLimit  As String
    
    IsCaseOwnerRemote = False 'default
    
    Set xmlReturnDoc = New FreeThreadedDOMDocument40
    xmlReturnDoc.validateOnParse = False
    xmlReturnDoc.setProperty "NewParser", True
    
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.validateOnParse = False
    xmlDoc.setProperty "NewParser", True
    
    Set xmlElem = xmlDoc.createElement("REQUEST")
    xmlElem.setAttribute "CRUD_OP", "READ"
    xmlElem.setAttribute "ENTITY_REF", "APPLICATION"
    xmlElem.setAttribute "SCHEMA_NAME", "TMRemoteOwnership"
    Set xmlNode = xmlDoc.appendChild(xmlElem)
    
    Set xmlApplicationNode = vxmlRequest.selectSingleNode("APPLICATION")
    
    Set xmlElem = xmlDoc.createElement("APPLICATION")
    xmlElem.setAttribute "APPLICATIONNUMBER", xmlGetAttributeText(xmlApplicationNode, "APPLICATIONNUMBER")
    xmlElem.setAttribute "APPLICATIONFACTFINDNUMBER", xmlGetAttributeText(xmlApplicationNode, "APPLICATIONFACTFINDNUMBER")
    
    ' ik_EP2_861_20070115
    If Len(xmlElem.getAttribute("APPLICATIONNUMBER")) = 0 Then
        Set xmlApplicationNode = vxmlRequest.selectSingleNode("CASETASK")
        xmlElem.setAttribute "APPLICATIONNUMBER", xmlGetAttributeText(xmlApplicationNode, "CASEID")
        xmlElem.setAttribute "APPLICATIONFACTFINDNUMBER", "1"
    End If
    
    xmlNode.appendChild xmlElem
    
    'Call CRUD to retrieve UserHistory data
    Set objCRUD = GetObjectContext.CreateInstance("omCRUD.omCRUDBO")
    
    If Not xmlReturnDoc.loadXML(objCRUD.omRequest(xmlDoc.xml)) Then
        errThrowError cstrFunctionName, oeXMLParserError, xmlReturnDoc.parseError.reason
    End If
    
    If xmlReturnDoc.selectSingleNode("RESPONSE[@TYPE='SUCCESS']") Is Nothing Then
        errThrowError cstrFunctionName, oeUnspecifiedError, xmlReturnDoc.xml
    End If
    
    If xmlReturnDoc.selectSingleNode("RESPONSE/APPLICATION/USERHISTORY") Is Nothing Then
        errThrowError cstrFunctionName, oeXMLMissingElement, "error retrieving USERHISTORY details"
    End If
    
    'If the UserRole record contains a role = "Remote Underwriter" (defined in CRUD schema file TMRemoteOwnership)
    Set xmlNode = xmlReturnDoc.selectSingleNode("RESPONSE/APPLICATION/USERHISTORY/USERROLE")
    If Not xmlNode Is Nothing Then
        'If USERROLEACTIVEFROM <= system date then check that Task.RemoteOwnerTaskInd exists
        If Len(xmlGetAttributeText(xmlNode, "USERROLEACTIVEFROM")) > 0 Then
            If CDate(xmlGetAttributeText(xmlNode, "USERROLEACTIVEFROM")) <= Now Then
                'AND UserRoleActiveTo > system date then
                'assign task ownership to remote owner
                If Len(xmlGetAttributeText(xmlNode, "USERROLEACTIVETO")) > 0 Then
                    If CDate(xmlGetAttributeText(xmlNode, "USERROLEACTIVETO")) > Now Then
                        rstrUnitID = xmlGetAttributeText(xmlNode, "UNITID")
                        rstrUserID = xmlGetAttributeText(xmlNode, "USERID")
                        blSuccess = True
                    End If
                'OR UserRoleActiveTo = NULL then
                'assign task ownership to remote owner
                Else
                    rstrUnitID = xmlGetAttributeText(xmlNode, "UNITID")
                    rstrUserID = xmlGetAttributeText(xmlNode, "USERID")
                    blSuccess = True
                End If
            End If
        End If
    End If
        
    IsCaseOwnerRemote = blSuccess
    
IsCaseOwnerRemote_Exit:
    
    Set xmlDoc = Nothing
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    Set xmlReturnDoc = Nothing
    Set objCRUD = Nothing
    Set xmlApplicationNode = Nothing
    Exit Function

IsCaseOwnerRemote_Error:
    
    Set xmlDoc = Nothing
    Set xmlElem = Nothing
    Set xmlNode = Nothing
    Set xmlReturnDoc = Nothing
    Set objCRUD = Nothing
    Set xmlApplicationNode = Nothing
    errCheckError cstrFunctionName
    
End Function
