VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "traceAssist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'TW     18/10/2005  MAR223
Private gblnTraceOn As Boolean
'TW 11/5/2005
Private gblnXMLTraceOn As Boolean
'TW 11/5/2005 End
Private gstrTraceFolder As String
Private gstrTraceID As String
Private gintStartElapsed As Single
Private gdtStartTime As Date
Public Sub StartTrace()
    Dim strRegValue As String, _
        strTempName As String
    Dim fso As Scripting.FileSystemObject
    Dim fsoFile As Scripting.TextStream
    Dim objWshShell As Object
    Set objWshShell = CreateObject("WScript.Shell")
    On Error Resume Next
    gdtStartTime = Now
    strRegValue = objWshShell.RegRead("HKLM\SOFTWARE\OMIGA4\TRACE\")
'TW 11/5/2005
'    If strRegValue = "1" Then
    If strRegValue = "1" Or strRegValue = "2" Then
        If strRegValue = "2" Then
            gblnXMLTraceOn = True
        End If
'TW 11/5/2005 End
        
        strRegValue = objWshShell.RegRead("HKLM\SOFTWARE\OMIGA4\TRACE\FOLDER")
        If Len(strRegValue) = 0 Then
            
            gblnTraceOn = False
            App.LogEvent App.Title & "tracing - no trace folder specified", vbLogEventTypeError
        Else
            
            Set fso = New Scripting.FileSystemObject
            If Not fso.FolderExists(strRegValue) Then
                
                gblnTraceOn = False
                App.LogEvent App.Title & "tracing - trace folder does not exist", vbLogEventTypeError
            Else
                
                gstrTraceFolder = strRegValue & "\"
                strTempName = fso.GetTempName
                strTempName = Left(strTempName, Len(strTempName) - 4)
                gstrTraceID = _
                    Format(gdtStartTime, "yyyymmdd_hhmmss") & "_" & App.ThreadID & "_" & strTempName
                gintStartElapsed = Timer
                gblnTraceOn = True
                    
            End If
        End If
    End If
    On Error GoTo 0
    Set fso = Nothing
    Set objWshShell = Nothing
End Sub
Public Sub TraceInitialiseFromParent( _
    ByVal vblnTraceOn As Boolean, _
    ByVal vstrFileName As String, _
    ByVal vintStartElapsed As Single)
    gblnTraceOn = vblnTraceOn
    gstrTraceID = vstrFileName
    gintStartElapsed = vintStartElapsed
End Sub
Public Sub TraceInitialiseOffspring(ByVal vobjOmComponent As Object)
    
    If Not gblnTraceOn Then
        Exit Sub
    End If
    vobjOmComponent.InitialiseTraceInterface gblnTraceOn, gstrTraceID, gintStartElapsed
End Sub
Public Sub TraceMethodEntry( _
    ByVal vstrModuleName As String, _
    ByVal vstrMethodName As String, _
    Optional ByVal vstrMessage As String)
    If Not gblnTraceOn Then
        Exit Sub
    End If
    TraceMessage vstrModuleName, vstrMethodName, "entry", vstrMessage
End Sub
Public Sub TraceMethodExit( _
    ByVal vstrModuleName As String, _
    ByVal vstrMethodName As String, _
    Optional ByVal vstrMessage As String)
    If Not gblnTraceOn Then
        Exit Sub
    End If
        
    If Err.Number <> 0 Then
        TraceMethodError vstrModuleName, vstrMethodName, vstrMessage
    Else
        TraceMessage vstrModuleName, vstrMethodName, "exit", vstrMessage
    End If
End Sub
Public Sub TraceMethodError( _
    ByVal vstrModuleName As String, _
    ByVal vstrMethodName As String, _
    Optional ByVal vstrMessage As String)
    Dim strErrMessage As String
    If Not gblnTraceOn Then
        Exit Sub
    End If
        
    If Len(vstrMessage) > 0 Then
        strErrMessage = vstrMessage & ", Error: "
    End If
    strErrMessage = _
        strErrMessage & _
        "number: " & Err.Number & _
        " description: " & Err.Description & _
        " source: " & Err.Source
    TraceMessage vstrModuleName, vstrMethodName, "error", strErrMessage
End Sub
Public Sub TraceMessage( _
    ByVal vstrModuleName As String, _
    ByVal vstrMethodName As String, _
    ByVal vstrMethodStage As String, _
    Optional ByVal vstrMessage As String)
    Dim fso As Scripting.FileSystemObject
    Dim fsoFile As Scripting.TextStream
    If Not gblnTraceOn Then
        Exit Sub
    End If
    Set fso = New Scripting.FileSystemObject
    If fso.FileExists(gstrTraceFolder & gstrTraceID & ".log") Then
        Set fsoFile = _
            fso.OpenTextFile(gstrTraceFolder & gstrTraceID & ".log", ForAppending, False)
        If fsoFile Is Nothing Then
            App.LogEvent App.Title & "tracing - cannot open trace file", vbLogEventTypeError
            Set fso = Nothing
            Exit Sub
        End If
            
    Else
        
        Set fsoFile = _
            fso.CreateTextFile(gstrTraceFolder & gstrTraceID & ".log")
        If fsoFile Is Nothing Then
            App.LogEvent App.Title & "tracing - cannot create trace file", vbLogEventTypeError
            Set fso = Nothing
            Exit Sub
        End If
        fsoFile.WriteLine "elapsed.component,module,method,stage,message"
    End If
        
    fsoFile.WriteLine _
        Format(Timer - gintStartElapsed, "0.000000") & "," & _
        App.Title & "," & _
        vstrModuleName & "," & _
        vstrMethodName & "," & _
        vstrMethodStage & "," & _
        vstrMessage
    fsoFile.Close
        
    Set fsoFile = Nothing
    Set fso = Nothing
End Sub
Public Sub TraceRequest(ByVal vstrXMLRequest As String)
    
    If Not gblnTraceOn Then
        Exit Sub
    End If
    TraceXML vstrXMLRequest, "request"
End Sub
Public Sub TraceResponse(ByVal vstrXmlResponse As String)
    
    If Not gblnTraceOn Then
        Exit Sub
    End If
    TraceXML vstrXmlResponse, "response"
End Sub
Public Sub TraceXML(ByVal vstrXML As String, ByVal vstrSuffix As String)
    
'TW 11/5/2005
'    If Not gblnTraceOn Then
'        Exit Sub
'    End If
    If Not gblnXMLTraceOn Then
        Exit Sub
    End If
'TW 11/5/2005 End
    Dim xmlDoc As FreeThreadedDOMDocument40
    Set xmlDoc = New FreeThreadedDOMDocument40
    Dim xmlElem As IXMLDOMElement
    xmlDoc.async = False
    xmlDoc.setProperty "NewParser", True
    xmlDoc.validateOnParse = False
    xmlDoc.LoadXML vstrXML
    If xmlDoc.parseError.errorCode <> 0 Then
        Set xmlElem = xmlDoc.createElement("XML_ERROR_TRACE")
        xmlElem.setAttribute "ERRORCODE", xmlDoc.parseError.errorCode
        xmlElem.setAttribute "FILEPOS", xmlDoc.parseError.filepos
        xmlElem.setAttribute "LINE", xmlDoc.parseError.Line
        xmlElem.setAttribute "LINEPOS", xmlDoc.parseError.linepos
        xmlElem.setAttribute "REASON", xmlDoc.parseError.reason
        xmlElem.Text = xmlDoc.parseError.srcText
        xmlDoc.appendChild xmlElem
    End If
    xmlDoc.save gstrTraceFolder & gstrTraceID & "_" & App.Title & "_" & vstrSuffix & ".xml"
    Set xmlElem = Nothing
    Set xmlDoc = Nothing
End Sub
Public Sub TraceIdentErrorResponse(ByRef vstrErrorResponse As String)
    
    If Not gblnTraceOn Then
        Exit Sub
    End If
    Dim xmlDoc As FreeThreadedDOMDocument40
    Set xmlDoc = New FreeThreadedDOMDocument40
    xmlDoc.async = False
    xmlDoc.setProperty "NewParser", True
    xmlDoc.validateOnParse = False
    xmlDoc.LoadXML vstrErrorResponse
    If xmlDoc.parseError.errorCode = 0 Then
        If Not xmlDoc.selectSingleNode("RESPONSE/ERROR/DESCRIPTION") Is Nothing Then
            If Left(xmlDoc.selectSingleNode("RESPONSE/ERROR/DESCRIPTION").Text, 11) <> "[trace id: " Then
                
                xmlDoc.selectSingleNode("RESPONSE/ERROR/DESCRIPTION").Text = _
                    "[trace id: " & gstrTraceID & "] " & _
                    xmlDoc.selectSingleNode("RESPONSE/ERROR/DESCRIPTION").Text
                vstrErrorResponse = xmlDoc.xml
            End If
        End If
    End If
    Set xmlDoc = Nothing
            
End Sub
