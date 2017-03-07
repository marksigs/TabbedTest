Attribute VB_Name = "AsyncAssist"
Option Explicit

Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
Private m_objRequest As IXMLDOMElement
Private m_objDataNode As IXMLDOMNode
Private m_lngTimerID As Long

Public Sub EnableAsyncDownload(ByVal vxmlRequest As IXMLDOMElement, ByVal xmlDataNode As IXMLDOMNode)
    
On Error GoTo EnableAsyncDownload_VbErr

    'Cache the Request
    Set m_objRequest = vxmlRequest
    Set m_objDataNode = xmlDataNode
    
    'Set the timer to start the download
    m_lngTimerID = SetTimer(0, 0, 100, AddressOf DoAsyncDownload)
    
    '... and return to the calling method
    Exit Sub

EnableAsyncDownload_VbErr:
    Err.Raise Err.Number, "[AsyncAssist].EnableAsyncDownload." & Err.Source, Err.Description
    
End Sub

Private Sub DoAsyncDownload(ByVal hWnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)

On Error GoTo DoAsyncDownload_VbErr

Dim objIDownload As IOmiga4ToOmiga3BO
Dim xmlTempResponseNode As IXMLDOMElement
Dim xmlDummyDoc As FreeThreadedDOMDocument40
Dim xmlDummyElem As IXMLDOMElement

    'Kill the timer - we do not need it any more
    KillTimer 0, m_lngTimerID
    
    Set xmlDummyDoc = New FreeThreadedDOMDocument40
    xmlDummyDoc.validateOnParse = False
    xmlDummyDoc.setProperty "NewParser", True
    Set objIDownload = New Omiga4toOmiga3BO
    
    Set xmlDummyElem = xmlDummyDoc.createElement("PERFORMANCE")
    Set xmlTempResponseNode = xmlDummyDoc.createElement("RESPONSE")
    xmlDummyDoc.appendChild xmlTempResponseNode
    xmlTempResponseNode.setAttribute "TYPE", "SUCCESS"

    #If ASYNCDEBUG = 1 Then
        Dim nFile As Integer
        nFile = FreeFile
        Open "C:\ASYNC_REQUEST.XML" For Output As #nFile
        Print #nFile, m_objRequest.XML
        Close #nFile
        Open "C:\ASYNC_DATA.XML" For Output As #nFile
        Print #nFile, m_objDataNode.XML
        Close #nFile
    #End If
        
    'Do the download
    objIDownload.CallTargetSystem m_objRequest, m_objDataNode, xmlTempResponseNode
    
    'Currently we do not check the response as this would require some
    'form of queue in Omiga4. For an Asynchronous Download assume success.

DoAsyncDownload_Exit:
    Set xmlTempResponseNode = Nothing
    Set objIDownload = Nothing
    Set xmlDummyElem = Nothing
    Set xmlDummyDoc = Nothing
    Exit Sub
    
DoAsyncDownload_VbErr:
    Err.Clear
    Exit Sub
    
End Sub

