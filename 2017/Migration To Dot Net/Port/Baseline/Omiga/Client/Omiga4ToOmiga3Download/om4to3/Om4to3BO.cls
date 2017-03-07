VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "Omiga4toOmiga3BO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      C:\Projects\dev\Omiga4ToOmiga3Download\om4to3\Om4to3BO.cls
' Copyright:    Copyright © 1999 Marlborough Stirling
'
'Description:
'   Processes a Download request for Omiga4 to Omiga3.
'   Recognises only two request types.


' Dependencies:
'   omBase

' Issues:
'------------------------------------------------------------------------------------------
'History:
'
'Prog  Date     Description
'MH    23/02/00 Created
'LDM   25/07/00 Amended test compression/decompression
'MC    15/08/00 SYS1411 - Add optional security data
'MC    16/08/00 Validate that a REQUEST tag is passed into AsyncDownload
'MC    07/09/00 SYS1411 - Add option to call om3BridgeInterface component
'CL    07/05/02 SYS4510  Modifications to class initialise
'CL    10/05/02 SYS4510  Remove class initialize & class terminate
'------------------------------------------------------------------------------------------
Option Explicit

Implements IOmiga4ToOmiga3BO
Implements ObjectControl

'=============================================
'Variable Declaration Section
'=============================================
Private m_objContext As ObjectContext
Private m_objErrAssist As ErrAssist
Private m_objXmlAssist As XMLAssist


Public Function Download(ByVal vstrXMLRequest As String) As String
' header ----------------------------------------------------------------------------------
' description:
'    Starts a download to Omiga3. Formats the download on the basis of
'    the Download type requested.
'
'   <REQUEST>
'     <DOWNLOAD>
'       <TYPE>Initial</TYPE>
'       <APPLICATIONNUMBER></APPLICATIONNUMBER>
'     </DOWNLOAD>
'   </REQUEST>
'
' pass:         vstrXmlRequest  The Application Number and the download type

' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------
On Error GoTo DownloadVbErr
    
    #If TIMINGS Then
        Dim t As Single
        Dim xmlTime As IXMLDOMElement
        t = Timer
    #End If

    Const strFunctionName As String = "Download"

        
    Dim xmlIn As FreeThreadedDOMDocument40
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
    Dim xmlDownloadResponseNode As IXMLDOMNode
    
    ' Create default response block
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
    
    #If TIMINGS Then
        Set xmlTime = xmlOut.createElement("PERFORMANCE")
        Set xmlTime = xmlOut.firstChild.appendChild(xmlTime)
        xmlTime.setAttribute "STARTTIME", Format$(Now, "hh:mm:ss")
    #End If
    
    Set xmlIn = m_objXmlAssist.load(vstrXMLRequest, TypeName(Me), strFunctionName)
    
    ' Delegate to DomDocument based method and attach returned data to our response
    #If TIMINGS Then
            Set xmlDownloadResponseNode = IOmiga4toOmiga3BO_Download(xmlIn.documentElement, xmlTime)
    #Else
            Set xmlDownloadResponseNode = IOmiga4toOmiga3BO_Download(xmlIn.documentElement)
    #End If
    
    m_objErrAssist.CheckResponse xmlDownloadResponseNode, xmlResponseElem, True
    
    m_objXmlAssist.AttachResponseData xmlResponseElem, xmlDownloadResponseNode
    
    #If TIMINGS Then
        xmlTime.setAttribute "ELAPSED", Format$(Timer - t, TIMEFORMAT)
        xmlDownloadResponseNode.appendChild xmlTime
    #End If
    
    Download = xmlDownloadResponseNode.XML
    
    
    If Not m_objContext Is Nothing Then
        m_objContext.SetComplete
    End If

DownloadExit:

    '   clean-up section, release object references
    Set xmlIn = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    Set xmlDownloadResponseNode = Nothing

    Exit Function

DownloadVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Download = m_objErrAssist.CreateErrorResponse

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DownloadExit

End Function

Private Sub IOmiga4ToOmiga3BO_AsyncDownload(ByVal vxmlRequest As IXMLDOMElement, ByVal xmlDataNode As IXMLDOMNode, xmlResponseElem As IXMLDOMElement, Optional ByVal vxmlPerformance As IXMLDOMElement)

On Error GoTo AsyncDownloadVbErr

    Const strFunctionName As String = "IOmiga4ToOmiga3BO_AsyncDownload"

    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    'Dim xmlResponseElem As IXMLDOMElement
    Dim xmlRequestNode As IXMLDOMNode
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"

    'Check that a REQUEST tag exists
    If vxmlRequest.nodeName = "REQUEST" Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName("REQUEST").Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "REQUEST tag not found"
    End If

    'This subroutine will call the download asynchronously
    EnableAsyncDownload vxmlRequest, xmlDataNode

    If Not m_objContext Is Nothing Then
        If xmlResponseElem.Attributes.getNamedItem("TYPE").Text = "SUCCESS" Then
            m_objContext.SetComplete
        Else
             m_objContext.SetAbort
        End If
    End If
    
AsyncDownloadExit:
    
    Set xmlOut = Nothing
    Set xmlRequestNode = Nothing
    
    Exit Sub

AsyncDownloadVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set xmlResponseElem = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume AsyncDownloadExit
    
End Sub

Private Sub IOmiga4ToOmiga3BO_CallTargetSystem(ByVal vxmlRequest As IXMLDOMElement, ByVal xmlDataNode As IXMLDOMNode, xmlResponseElem As IXMLDOMElement)

On Error GoTo CallTargetSystemVbErr

    Const strFunctionName As String = "IOmiga4ToOmiga3BO_CallTargetSystem"

    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim strResponse As String
    Dim objOmiga3Download As Object
    Const OM3BRIDGEINTERFACE = "OM3BRIDGEINTERFACE.OM3BRIDGEINTERFACE1"
    
    #If TIMINGS Then
        Dim t As Single
        t = Timer
    #End If
    
    Dim xmlOutputNode As IXMLDOMElement
    Set xmlOutputNode = vxmlRequest.getElementsByTagName("TARGETSYSTEM").Item(0)
    
    If xmlOutputNode Is Nothing Then
        'Run with Omiga3 as the default
        xmlResponseElem.setAttribute "TARGETSYSTEM", "OM3MANAGER.CUSTREGBO.Create()"
        If Not m_objContext Is Nothing Then
            Set objOmiga3Download = m_objContext.CreateInstance(gstrOMIGA3_MANAGER_COMPONENT & ".CustRegBo")
        Else
            Set objOmiga3Download = CreateObject(gstrOMIGA3_MANAGER_COMPONENT & ".CustRegBo")
        End If
    
       strResponse = objOmiga3Download.Create(xmlDataNode.XML)
    Else
        Dim strBO As String
        Dim strMethod As String
        strBO = SafeNullString(xmlOutputNode.getAttribute("BO"))
        strMethod = SafeNullString(xmlOutputNode.getAttribute("METHOD"))
        
        If strBO = vbNullString And strMethod = vbNullString Then
            xmlResponseElem.setAttribute "TARGETSYSTEM", "Omiga3Download not requested"
        ElseIf strBO = "" Or strMethod = "" Then
            m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingPrimaryTag, _
                    xmlOutputNode.XML & " BO or METHOD attribute is Null"
        'SYS1411 MDC 06/09/2000. Call OM3BRIDGEINTERFACE component
        ElseIf UCase$(Trim$(strBO)) = OM3BRIDGEINTERFACE Then
            Dim lngDownloadType As Long
            Dim xmlDownloadTypeNode As IXMLDOMElement
            Set xmlDownloadTypeNode = vxmlRequest.getElementsByTagName("TYPE").Item(0)
            If xmlDownloadTypeNode Is Nothing Then
                m_objErrAssist.ThrowError TypeName(Me), strFunctionName, oeMissingElement, _
                    "An element with the name 'TYPE' was expected but not found"
            Else
                Select Case UCase$(Trim$(xmlDownloadTypeNode.Text))
                Case "AIP"
                    lngDownloadType = 0
                Case "FORMFILL"
                    lngDownloadType = 1
                Case Else
                    'Assume AIP
                    lngDownloadType = 0
                End Select
            End If
            Set xmlDownloadTypeNode = Nothing
            xmlResponseElem.setAttribute "TARGETSYSTEM", strBO & "." & strMethod & "()"
            Set objOmiga3Download = CreateObject(OM3BRIDGEINTERFACE)
            objOmiga3Download.Run lngDownloadType, xmlDataNode.XML, strResponse
        Else
            xmlResponseElem.setAttribute "TARGETSYSTEM", strBO & "." & strMethod & "()"
            If Not m_objContext Is Nothing Then
                Set objOmiga3Download = m_objContext.CreateInstance(strBO)
            Else
                Set objOmiga3Download = CreateObject(strBO)
            End If
        
           'Now make that call
           strResponse = CallByName(objOmiga3Download, strMethod, VbMethod, xmlDataNode.XML)
        End If
    End If
    
    xmlOut.async = True
    xmlOut.loadXML strResponse
    Set xmlResponseElem = xmlOut.getElementsByTagName("RESPONSE").Item(0)
                            
CallTargetSystemExit:
    
    Set xmlOut = Nothing
    
    Exit Sub

CallTargetSystemVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set xmlResponseElem = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume CallTargetSystemExit

End Sub

Private Function IOmiga4toOmiga3BO_Download(ByVal vxmlRequest As IXMLDOMElement, _
                    Optional ByVal vxmlPerformance As IXMLDOMElement) As IXMLDOMNode
' header ----------------------------------------------------------------------------------
' description:
'    Starts a download to Omiga3. Formats the download on the basis of
'    the Download type requested.
'
' pass:         vstrXmlRequest  The Application Number and the download type
'   <REQUEST ACTION="DOWNLOAD">
'       <TYPE>AIP</TYPE> or <TYPE>FormFill</TYPE>
'       <APPLICATIONNUMBER></APPLICATIONNUMBER>
'     </DOWNLOAD>
'   </REQUEST>
'

' return:                       xml Response data stream containing results of operation
'                               either: TYPE="SUCCESS"
'                               or: TYPE="SYSERR" and <ERROR> element
'------------------------------------------------------------------------------------------

On Error GoTo DownloadVbErr
    
    Const strFunctionName As String = "IOmiga4toOmiga3BO_Download"


    Dim xmlRequestNode As IXMLDOMNode
    Dim xmlRequestElem As IXMLDOMElement
    Dim xmlDataNode As IXMLDOMNode
    Dim xmlOut As New FreeThreadedDOMDocument40
    xmlOut.validateOnParse = False
    xmlOut.setProperty "NewParser", True
    Dim xmlResponseElem As IXMLDOMElement
        
    Dim objIOmiga4DO As IOmiga4toOmiga3DO
    Dim objOmiga3Download As Object
    Dim strResponse As String
    
    Dim xmlNode As IXMLDOMNode
    Dim MyXMLDoc As FreeThreadedDOMDocument40
    Dim xmlSecurity As IXMLDOMNode
    
    Set xmlResponseElem = xmlOut.createElement("RESPONSE")
    xmlOut.appendChild xmlResponseElem
    xmlResponseElem.setAttribute "TYPE", "SUCCESS"
      
    If m_objContext Is Nothing Then
        Set objIOmiga4DO = New Omiga4toOmiga3DO
    Else
        Set objIOmiga4DO = m_objContext.CreateInstance(App.Title & ".Omiga4toOmiga3DO")
    End If
    
    If vxmlRequest.nodeName = "REQUEST" Then
        Set xmlRequestNode = vxmlRequest
    Else
        Set xmlRequestNode = vxmlRequest.getElementsByTagName("REQUEST").Item(0)
    End If
        
    If xmlRequestNode Is Nothing Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "REQUEST tag not found"
    End If
        
    Set xmlRequestElem = xmlRequestNode
    
    If IsNull(xmlRequestElem.getAttribute("ACTION")) Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "ACTION attribute tag not found"
    End If
    
    If UCase$(Trim$(xmlRequestElem.getAttribute("ACTION"))) <> "DOWNLOAD" Then
        m_objErrAssist.ThrowError TypeName(Me), strFunctionName, _
            oeMissingPrimaryTag, "Expected attribute tag ACTION=""DOWNLOAD"", received ACTION=""" & xmlRequestElem.getAttribute("ACTION") & """ on REQUEST Elem"
    End If
           
    xmlResponseElem.setAttribute ("APPLICATIONNUMBER"), m_objXmlAssist.GetTagValue(xmlRequestElem, "APPLICATIONNUMBER", False)
    xmlResponseElem.setAttribute ("DOWNLOADTYPE"), m_objXmlAssist.GetTagValue(xmlRequestElem, "TYPE", False)
    
    #If TIMINGS Then
        vxmlPerformance.setAttribute "APPLICATIONNUMBER", xmlResponseElem.getAttribute("APPLICATIONNUMBER")
        vxmlPerformance.setAttribute "DOWNLOADTYPE", xmlResponseElem.getAttribute("DOWNLOADTYPE")

        Set xmlDataNode = objIOmiga4DO.GetData(xmlRequestNode, vxmlPerformance)
    #Else
        Set xmlDataNode = objIOmiga4DO.GetData(xmlRequestNode)
    #End If
    
    'SYS1411 - Append optional password data to the downloaded XML
    Set xmlSecurity = vxmlRequest.selectSingleNode(".//SECURITYDETAILS")
    If Not xmlSecurity Is Nothing Then
        Set MyXMLDoc = New FreeThreadedDOMDocument40
        MyXMLDoc.validateOnParse = False
        MyXMLDoc.setProperty "NewParser", True
        MyXMLDoc.loadXML xmlDataNode.XML
        Set xmlDataNode = MyXMLDoc.selectSingleNode(".//REQUEST")
        xmlDataNode.appendChild xmlSecurity
    End If
    
    Dim strFileName As String
    strFileName = SafeNullString(vxmlRequest.getAttribute("OUTPUTFILE"))
    If strFileName <> vbNullString Then
        xmlResponseElem.setAttribute "OUTPUTFILE", strFileName
        Open strFileName For Output As 1 Len = 8192
        Print #1, xmlDataNode.XML
        Close 1
    End If

    #If NODOWNLOAD Then
        xmlResponseElem.setAttribute "COMMENT", "OMIGA3 DOWNLOAD COMMENTED OUT"
    #Else
        'Check if download to be run asynchronously
        Dim strAsync As String
        strAsync = m_objXmlAssist.GetAttributeValue(vxmlRequest, "REQUEST", "ASYNC")
        If Len(Trim$(strAsync)) > 0 Then
            'Call Download asynchronously
            IOmiga4ToOmiga3BO_AsyncDownload vxmlRequest, xmlDataNode, xmlResponseElem
        Else
            'Call Download synchronously
            IOmiga4ToOmiga3BO_CallTargetSystem vxmlRequest, xmlDataNode, xmlResponseElem
        End If
         
    #End If
   
    #If TIMINGS Then
        vxmlPerformance.setAttribute "XMLTOOM3", Format$(Timer - t, TIMEFORMAT)
    #End If

    #If TIMINGS Then
        vxmlPerformance.setAttribute "XMLBYTES", Len(xmlDataNode.XML)
        vxmlPerformance.setAttribute "TYPE", xmlResponseElem.getAttribute("TYPE")
    #End If
    
        
    Set IOmiga4toOmiga3BO_Download = xmlResponseElem
    
    If Not m_objContext Is Nothing Then
        If xmlResponseElem.Attributes.getNamedItem("TYPE").Text = "SUCCESS" Then
            m_objContext.SetComplete
        Else
             m_objContext.SetAbort
        End If
    End If
    
DownloadExit:
    
    Set xmlRequestNode = Nothing
    Set objIOmiga4DO = Nothing
    Set xmlDataNode = Nothing
    Set xmlOut = Nothing
    Set xmlResponseElem = Nothing
    
    Set xmlNode = Nothing
    Set MyXMLDoc = Nothing
    Set xmlSecurity = Nothing
    
    Exit Function

DownloadVbErr:

    If m_objErrAssist.IsWarning = True Then
        m_objErrAssist.AddWarning xmlResponseElem
        Resume Next
    End If

    m_objErrAssist.AddToErrSource strFunctionName
    
    If m_objErrAssist.IsSystemError = True Then
        m_objErrAssist.LogError TypeName(Me), strFunctionName, Err.Number, Err.Description
    End If

    Set IOmiga4toOmiga3BO_Download = m_objErrAssist.CreateErrorResponseEx()

    If Not m_objContext Is Nothing Then
        m_objContext.SetAbort
    End If
    
    ' go to clean-up section
    Resume DownloadExit

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

Private Function SafeNullString(ByVal vvarValue As Variant) As String
    If IsNull(vvarValue) Then
        SafeNullString = vbNullString
    Else
        SafeNullString = vvarValue
    End If
End Function

