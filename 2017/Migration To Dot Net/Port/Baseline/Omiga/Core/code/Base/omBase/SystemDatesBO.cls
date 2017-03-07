VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "SystemDatesBO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      SystemDatesBO.cls
'Copyright:     Copyright © 1999 Marlborough Stirling
'Description:   Contains methods that perform date manipulation with regard to working and
'               non-working days in relation to nonworking days specified at country and
'               distribution channel level.
'
'Dependencies:
'               SystemDatesDO
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date        Description
'DJP    17/01/01    Created
'DJP    23/01/01    Added ValidateTimePart to validate days and hours
'DJP    29/01/01    In FindWorkingDay move 0 hours onto next day.
'CL     07/05/02    SYS4510  Modifications to class initialise
'CL     10/05/02    SYS4510  Remove class initialize & class terminate
'------------------------------------------------------------------------------------------
Option Explicit
Implements ISystemDatesBO
Implements ObjectControl
'=============================================
'Constant Declaration Section
Private Const cstrROOT_NODE_NAME = "SYSTEMDATE"
'=============================================
'Variable Declaration Section
'=============================================
Private m_objContext As ObjectContext
Private m_objErrAssist As ErrAssist
Private m_objXmlAssist As XMLAssist
Public Function FindWorkingDay(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  See ISystemDatesBO_FindWorkingDay
' pass:         vstrXmlRequest  xml Request data stream
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo FindWorkingDayVbErr
    Const strFunctionName As String = "FindWorkingDay"
        
        
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                    
    ' Delegate to FreeThreadedDOMDocument40 based method and attach returned data to our response
    Set xmlTempResponseNode = ISystemDatesBO_FindWorkingDay(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    FindWorkingDay = xmlResponseElem.xml
    m_objContext.SetComplete
FindWorkingDayExit:
    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Exit Function
FindWorkingDayVbErr:
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    FindWorkingDay = m_objErrAssist.CreateErrorResponse
    m_objContext.SetComplete
    ' go to clean-up section
    Resume FindWorkingDayExit
End Function
Private Function ISystemDatesBO_FindWorkingDay(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Returns the next or previous working day, or if variation is specified, finds the n'th
'               non working day either before or after the specified date.
'               business object
' pass:         vxmlRequest  xml Request data stream containing data to be persisted
'               DATE: The start date and optionally, time
'               DIRECTION: either + or - for future or past
'               HOURS: number of hours to move forwards or backwards
'               DAYS: number of days to move forwards or backwards
'               CHANNELID: distribution channel for which nonworking days should be found
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo FindWorkingDayVbErr
    
    Const strFunctionName As String = "ISystemDatesBO_FindWorkingDay"
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlNonWorking As FreeThreadedDOMDocument40
    Dim objISystemDatesDO As ISystemDatesDO
    Dim xmlComboResp As FreeThreadedDOMDocument40
    Dim xmlChannelID As IXMLDOMElement
    Set xmlOut = New FreeThreadedDOMDocument40
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    Set xmlResponseElem = xmlResponseElem.appendChild(xmlOut.createElement(cstrROOT_NODE_NAME))
    If vxmlRequest.nodeName = cstrROOT_NODE_NAME Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName(cstrROOT_NODE_NAME).Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrROOT_NODE_NAME & " tag not found"
    End If
    Set objISystemDatesDO = m_objContext.CreateInstance(App.Title & ".SystemDatesDO")
    Dim strDate As Date
    Dim strInputDate As String
    Dim strTmpDate As Date
    Dim strDirection As String
    Dim strHours As String
    Dim strDays As String
    Dim strDistChannel As String
    Dim bLimit As Boolean
    Dim bFound As Boolean
    Dim strNonWorkingDay As String
    Dim xmlDateOccurence As FreeThreadedDOMDocument40
    Dim xmlDateElem As IXMLDOMElement
    Dim lngLimit As Long
    Dim lngDayCount As Long
    Dim xmlWorkingDayResp As IXMLDOMElement
    bLimit = False
    ' Read input XML Values
    strInputDate = m_objXmlAssist.GetNodeValue(xmlRequestNode, ".//DATE")
    strDirection = m_objXmlAssist.GetNodeValue(xmlRequestNode, ".//DIRECTION")
    ' Both the following are non-mandatory
    strHours = m_objXmlAssist.GetNodeValue(xmlRequestNode, ".//HOURS", False)
    strDays = m_objXmlAssist.GetNodeValue(xmlRequestNode, ".//DAYS", False)
    Set xmlChannelID = m_objXmlAssist.GetNode(xmlRequestNode, ".//CHANNELID").cloneNode(True)
    ' Validation direction
    If strDirection <> "+" And strDirection <> "-" Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeInvalidParameter, "Direction must be '+' or '-'"
    End If
    ' Validate Hours
    If Len(strDays) > 0 Then
        bLimit = True
        If Not ValidateTimePart(strDays) Then
            m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeInvalidParameter, "DAYS must be a positive whole number"
        End If
    End If
    ' Validate Minutes
    If Len(strHours) > 0 Then
        If Not ValidateTimePart(strHours) Then
            m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeInvalidParameter, "HOURS must be a positive whole number"
        End If
    End If
    ' Validate Date
    If Not IsDate(strInputDate) Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
        oeInvalidParameter, "DATE must be a validate date or date/time"
    End If
    strDate = CDate(strInputDate)
    ' Need to loop from the data passed in to the specified number of days and hours, or if days and hours
    ' are not specified, to the next non working day.
    bFound = False
    ' Create request block for SystemDatesDO
    Set xmlDateOccurence = New FreeThreadedDOMDocument40
    Set xmlDateElem = xmlDateOccurence.appendChild(xmlDateOccurence.createElement(cstrROOT_NODE_NAME))
    xmlDateElem.appendChild xmlChannelID
    Set xmlDateElem = xmlDateElem.appendChild(xmlDateOccurence.createElement("DATE"))
    lngLimit = 0
    lngDayCount = 0
    ' Decide if we're going forwards or backwards
    If strDirection = "-" Then
        lngLimit = -1
    Else
        lngLimit = 1
    End If
    strTmpDate = strDate
    ' Add the hours first
    If Len(strHours) > 0 Then
        If IsNumeric(strHours) Then
            
            If strDirection = "-" Then
                strDate = strDate - (1 / 24 * strHours)
            Else
                strDate = strDate + (1 / 24 * strHours)
            End If
        Else
            m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
                oeInvalidParameter, "Hours must be a single integer value"
        End If
    End If
    ' If we've not jumped over a day boundry, start on the next day
    If Day(strTmpDate) = Day(strDate) Then
        lngDayCount = 1
        ' Add or subtract a day to the current date for our start point. If adding the hour onto
        ' the date did this, we won't come in here so won't add the date
        strDate = strDate + lngLimit
    End If
    Do While bFound = False
        
        ' See if this day is a non working day or not
        xmlDateElem.Text = strDate
        ' Is the day a working day?
        Set xmlWorkingDayResp = objISystemDatesDO.CheckNonWorkingOccurence(xmlDateOccurence.documentElement)
        strNonWorkingDay = m_objXmlAssist.GetNodeValue(xmlWorkingDayResp, "//NONWORKINGIND")
        ' Decide if we've hit a non-working day or not. If we have and there isn't a day count specified we've
        ' found the right day. If not, we need to count up until the specified number of days has been hit.
        If strNonWorkingDay = "0" Then
            If bLimit Then
                If lngDayCount = CLng(strDays) Then
                    bFound = True
                Else
                    lngDayCount = lngDayCount + 1
                End If
            Else
                bFound = True
            End If
        End If
        If Not bFound Then
            strDate = strDate + lngLimit
        End If
    Loop
    Set xmlDataNode = xmlResponseElem.appendChild(xmlOut.createElement("DATE"))
    xmlDataNode.Text = strDate
        
    Set ISystemDatesBO_FindWorkingDay = xmlOut.documentElement
    m_objContext.SetComplete
FindWorkingDayExit:
    
    Set xmlRequestNode = Nothing
    Set objISystemDatesDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Exit Function
FindWorkingDayVbErr:
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    Set ISystemDatesBO_FindWorkingDay = m_objErrAssist.CreateErrorResponseEx()
    m_objContext.SetComplete
    ' go to clean-up section
    Resume FindWorkingDayExit
End Function
Private Function ValidateTimePart(strNumber As String)
' header ----------------------------------------------------------------------------------
' description:  ValidateTimePart - Validates that the number passed in is a) there at all,
'               b) a whole number and c) positive
' pass:         strNumber       Number to be validated
' return:       true if valid, false if not
'------------------------------------------------------------------------------------------
    Dim bRet As Boolean
    Dim dNumber As Double
    Dim strFunctionName As String
    On Error GoTo ValidateTimePartVbErr
    strFunctionName = "ValidateTimePart"
    bRet = False
    If Len(strNumber) > 0 Then
        If IsNumeric(strNumber) Then
            dNumber = CDbl(strNumber)
            If dNumber >= 0 Then
                If Fix(dNumber) - dNumber = 0 Then
                    bRet = True
                End If
            End If
        End If
    End If
    ValidateTimePart = bRet
    Exit Function
ValidateTimePartVbErr:
        m_objErrAssist.AddToErrSource strFunctionName
        Err.Raise Err.Number, Err.Source, Err.Description
End Function
Private Function ISystemDatesBO_CheckNonWorkingOccurence(ByVal vxmlRequest As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:  Returns whether or not the date passed in is a non-working day or not.
' pass:         vxmlRequest  xml Request data stream
' return:       xml Response Node
'------------------------------------------------------------------------------------------
On Error GoTo CheckNonWorkingOccurenceVbErr
    
    Const strFunctionName As String = "ISystemDatesBO_CheckNonWorkingOccurence"
    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim objISystemDatesDO As ISystemDatesDO
    Dim xmlNonWorking As FreeThreadedDOMDocument40
    Set xmlOut = New FreeThreadedDOMDocument40
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    Set xmlResponseElem = xmlResponseElem.appendChild(xmlOut.createElement(cstrROOT_NODE_NAME))
    Set objISystemDatesDO = m_objContext.CreateInstance(App.Title & ".SystemDatesDO")
    If vxmlRequest.nodeName = cstrROOT_NODE_NAME Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName(cstrROOT_NODE_NAME).Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, cstrROOT_NODE_NAME & " tag not found"
    End If
        
    ' Call the DO...
    Dim xmlNonWorkingelem As IXMLDOMElement
    Set xmlNonWorking = New FreeThreadedDOMDocument40
    Set xmlNonWorkingelem = xmlNonWorking.appendChild(xmlRequestNode.cloneNode(True))
    Set xmlDataNode = objISystemDatesDO.CheckNonWorkingOccurence(xmlNonWorking.documentElement)
    xmlResponseElem.appendChild xmlDataNode
        
    Set ISystemDatesBO_CheckNonWorkingOccurence = xmlOut.documentElement
    m_objContext.SetComplete
CheckNonWorkingOccurenceExit:
    
    Set xmlRequestNode = Nothing
    Set objISystemDatesDO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Exit Function
CheckNonWorkingOccurenceVbErr:
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    Set ISystemDatesBO_CheckNonWorkingOccurence = m_objErrAssist.CreateErrorResponseEx()
    m_objContext.SetComplete
    ' go to clean-up section
    Resume CheckNonWorkingOccurenceExit
End Function
Public Function CheckNonWorkingOccurence(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:  See ISystemDatesBO_CheckNonWorkingOccurence
' pass:         vstrXmlRequest  xml Request data stream containing data to be persisted
' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo CheckNonWorkingOccurenceVbErr
    Const strFunctionName As String = "CheckNonWorkingOccurence"
        
        
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As FreeThreadedDOMDocument40
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlTempResponseNode As IXMLDOMNode
    ' Create default response block
    Set xmlOut = New FreeThreadedDOMDocument40
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
                                    
    ' Delegate to FreeThreadedDOMDocument40 based method and attach returned data to our response
    Set xmlTempResponseNode = ISystemDatesBO_CheckNonWorkingOccurence(xmlIn.documentElement)
    m_objErrAssist.CheckResponse xmlTempResponseNode, xmlResponseElem, True
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlTempResponseNode
    CheckNonWorkingOccurence = xmlOut.xml
    m_objContext.SetComplete
CheckNonWorkingOccurenceExit:
    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlTempResponseNode = Nothing
    Exit Function
CheckNonWorkingOccurenceVbErr:
    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If
    m_objErrAssist.AddToErrSource strFunctionName
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If
    CheckNonWorkingOccurence = m_objErrAssist.CreateErrorResponse
    m_objContext.SetComplete
    ' go to clean-up section
    Resume CheckNonWorkingOccurenceExit
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
