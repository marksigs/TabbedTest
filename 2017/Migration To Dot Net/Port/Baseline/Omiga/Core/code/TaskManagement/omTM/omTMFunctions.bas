Attribute VB_Name = "omTMFunctions"
'Workfile:      CommonFunctions.bas
'Copyright:     Copyright © 2005 Marlborough Stirling
'Description:   Functions common to other modules.
'
'-------------------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'AS     25/07/2005  COER174 Created by extracting AddDefaultValuesToCaseTask from omTmNoTxBO.cls and omTMBO.cls.
'------------------------------------------------------------------------------------------------------

Option Explicit

Public Sub AddDefaultValuesToCaseTask( _
    ByRef objContext As ObjectContext, _
    ByVal vxmlRequestNode As IXMLDOMNode, _
    ByVal vxmlCaseTaskNode As IXMLDOMNode, _
    ByVal vxmlTaskNode As IXMLDOMNode, _
    ByVal vstrDefaultUserId As String, _
    ByVal vStrDefaultUnitId As String)
    
    On Error GoTo AddDefaultValuesToCaseTaskExit
    
    Dim strFunctionName As String
    strFunctionName = "AddDefaultValuesToCaseTask"
    
    Dim strTaskUserID As String
    Dim strUnitId As String
    Dim strTaskOwnerType As String
    Dim strTmpResponse As String
    Dim strOwningUserID As String
    Dim strOwningUnitID As String
    Dim objOrg As Object
      
    Dim xmlIn As New FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    Dim xmlElement As IXMLDOMElement
    Dim xmlTempNode As IXMLDOMNode
    Dim xmlParentNode As IXMLDOMNode
    Dim xmlChildNode As IXMLDOMNode
    Dim lngErrNo As Long
    Set xmlElement = vxmlRequestNode
    Set xmlParentNode = xmlIn.appendChild(xmlElement)
    If xmlAttributeValueExists(vxmlCaseTaskNode, "CASETASKNAME") = False Then
        xmlCopyAttributeValue vxmlTaskNode, vxmlCaseTaskNode, "TASKNAME", "CASETASKNAME"
    End If
            
    If xmlAttributeValueExists(vxmlCaseTaskNode, "TASKDUEDATEANDTIME") = False Then
        SetCaseTaskDueDateAndTime vxmlCaseTaskNode, vxmlTaskNode
    End If
    
    '  Getting the OwningUnitID and OwningUserID
    'GD BM0340 START
    strTaskUserID = xmlGetAttributeText(vxmlTaskNode, "TASKUSERID")
    'GD BM0340 END
    
    Dim strCaseTaskUserID As String
    Dim strCaseTaskUnitID As String
    strCaseTaskUserID = xmlGetAttributeText(vxmlCaseTaskNode, "OWNINGUSERID")
    strCaseTaskUnitID = xmlGetAttributeText(vxmlCaseTaskNode, "OWNINGUNITID")
    
    If strCaseTaskUserID <> "" Then
        'AS 25/05/2005 CORE174 Use OWNINGUSERID and OWNINGUNITID in request.
        'This enables the user to create an ad hoc task for a nominated user.
        strOwningUserID = strCaseTaskUserID
        strOwningUnitID = strCaseTaskUnitID
        
    ElseIf strTaskUserID <> "" Then
        'GD BM0340 START
        strOwningUserID = xmlGetAttributeText(vxmlTaskNode, "TASKUSERID")
        strOwningUnitID = xmlGetAttributeText(vxmlTaskNode, "TASKUNITID")
        'GD BM0340 END
        
    Else 'Non Stage Task User
        strTaskOwnerType = xmlGetAttributeText(vxmlTaskNode, "TASKOWNERTYPE")
        If strTaskOwnerType <> "" Then
            
            'Build up request for call to omOrg.OrganisationBO.GetActionOwnerDetails
            
            'GD BM0340 START
            strUnitId = xmlGetAttributeText(vxmlTaskNode, "TASKUNITID")
            'GD BM0340 END
            If strUnitId = "" Then 'Stage Task Unit undefined
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
            Set xmlOut = xmlLoad(strTmpResponse, strFunctionName)
            
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
                    errCheckXMLResponseNode xmlOut, , True
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
    xmlSetAttributeValue vxmlCaseTaskNode, "OWNINGUSERID", strOwningUserID
    xmlSetAttributeValue vxmlCaseTaskNode, "OWNINGUNITID", strOwningUnitID
    'BMIDS01112 MDC 29/11/2002 - Ensure that Mandatory tasks are marked Mandatory
    If Not xmlAttributeValueExists(vxmlCaseTaskNode, "MANDATORYINDICATOR") Then
        xmlCopyAttributeValue vxmlTaskNode, vxmlCaseTaskNode, "MANDATORYFLAG", "MANDATORYINDICATOR"
    End If
    'BMIDS01112 MDC 29/11/2002 - End
    
AddDefaultValuesToCaseTaskExit:
    
    Set objOrg = Nothing
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlTempNode = Nothing
    Set xmlElement = Nothing
    Set xmlParentNode = Nothing
    Set xmlChildNode = Nothing
    errCheckError strFunctionName
End Sub

Public Sub SetCaseTaskDueDateAndTime( _
    ByVal vxmlCaseTaskNode As IXMLDOMNode, _
    ByVal vxmlTaskNode As IXMLDOMNode)
    Const cstrFunctionName As String = "SetCaseTaskDueDateAndTime"
    On Error GoTo SetCaseTaskDueDateAndTimeExit
    Dim dtDueDateTime As Date
    Dim intChasePeriodDays As Integer, _
        intChasePeriodHours As Integer, _
        intOffsetDays As Integer, _
        intOffsetHours As Integer
    dtDueDateTime = Now()
    intChasePeriodDays = xmlGetAttributeAsInteger(vxmlTaskNode, "CHASINGPERIODDAYS")
    intChasePeriodHours = xmlGetAttributeAsInteger(vxmlTaskNode, "CHASINGPERIODHOURS")
    If intChasePeriodDays > 0 Then
        dtDueDateTime = DateAdd("y", intChasePeriodDays, dtDueDateTime)
    End If
    If intChasePeriodHours > 0 Then
        dtDueDateTime = DateAdd("h", intChasePeriodHours, dtDueDateTime)
    End If
    intOffsetDays = xmlGetAttributeAsInteger(vxmlTaskNode, "ADJUSTMENTDAYS")
    intOffsetHours = xmlGetAttributeAsInteger(vxmlTaskNode, "ADJUSTMENTHOURS")
    If intOffsetDays <> 0 Then
        dtDueDateTime = DateAdd("y", intOffsetDays, dtDueDateTime)
    End If
    If intOffsetHours <> 0 Then
        dtDueDateTime = DateAdd("h", intOffsetHours, dtDueDateTime)
    End If
    If dtDueDateTime < Now() Then
        dtDueDateTime = Now()
    End If
    xmlSetAttributeValue vxmlCaseTaskNode, "TASKDUEDATEANDTIME", dtDueDateTime
SetCaseTaskDueDateAndTimeExit:

    errCheckError cstrFunctionName
End Sub


