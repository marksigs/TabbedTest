VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Log"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------------
'Workfile:      Log.cls
'Copyright:     Copyright © 2005 Marlborough Stirling

'Description:   Log to file.
'Dependencies:  None
'Issues:
'
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date     Description
'AS     15/08/05 Created.
'------------------------------------------------------------------------------------------

Option Explicit

Private m_fileName As String
Private m_hFile As Integer
Private m_showError As Boolean
Private m_enableLog As Boolean
Private m_xmlFileIndex As Long

Private Sub Class_Initialize()
    m_fileName = ""
    m_hFile = 0
    m_showError = True
    m_enableLog = False
    m_xmlFileIndex = 0
End Sub

Private Sub Class_Terminate()
    CloseLog
End Sub

Public Property Get FileName() As String
    FileName = m_fileName
End Property

Public Property Let FileName(ByVal fName As String)
    m_fileName = fName
End Property

Public Property Get EnableLog() As Boolean
    EnableLog = m_enableLog
End Property

Public Property Let EnableLog(ByVal enable As Boolean)
    m_enableLog = enable
End Property

Public Function OpenLog(Optional ByVal fName As String = "", Optional ByVal asNew As Boolean = True) As Boolean
On Error GoTo errorLabel
    
    If m_hFile = 0 And m_enableLog Then
        OpenLog = False
        
        If fName <> "" Then
            m_fileName = fName
        End If
        
        If m_fileName = "" Then
            m_fileName = App.Path & "\" & App.EXEName & ".log"
        End If
        
        m_hFile = FreeFile()
        If asNew Then
            Open m_fileName For Output Access Write Lock Write As m_hFile
            WriteAbout
        Else
            Open m_fileName For Append Access Write Lock Write As m_hFile
        End If
                
        OpenLog = True
    Else
        ' Log is already open; do not open again.
        OpenLog = True
    End If
    
exitLabel:
    Exit Function
    
errorLabel:
    ReportError Err, "Log.OpenLog", "Opening file " & m_fileName
    GoTo exitLabel
    
End Function

Public Function CloseLog() As Boolean
On Error GoTo errorLabel
    
    CloseLog = False
    
    If m_enableLog Then
        If m_hFile <> 0 Then
            Close m_hFile
        End If
        m_hFile = 0
    End If
    
    CloseLog = True
    
exitLabel:
    Exit Function

errorLabel:
    ReportError Err, "Log.CloseLog", "Closing file " & m_fileName
    GoTo exitLabel
    
End Function

Public Function WriteLine(ByVal line As String) As Boolean
On Error GoTo errorLabel

    WriteLine = False
    
    If m_enableLog Then
        If OpenLog(, False) Then
            Dim dtNow As Date
            
            dtNow = Now
            
            Print #m_hFile, dtNow & ": " & line
            CloseLog    ' Required to flush the file.
            
        End If
    End If
    
    WriteLine = True
    
exitLabel:
    Exit Function

errorLabel:
    ReportError Err, "Log.WriteLine", "Writing to file " & m_fileName
    GoTo exitLabel
    
End Function

Public Function WriteLineXml(ByVal line As String, ByVal xml As String) As Boolean
On Error GoTo errorLabel

    WriteLineXml = False
    
    If m_enableLog Then
        WriteLine line & "(" & WriteXml(xml) & ")"
    End If
    
    WriteLineXml = True
    
exitLabel:
    Exit Function

errorLabel:
    ReportError Err, "Log.WriteLineXml", "Writing to file " & m_fileName
    GoTo exitLabel
    
End Function

Public Function WriteXml(ByVal xml As String) As String
On Error GoTo errorLabel

    Dim xmlDoc As MSXML2.DOMDocument
    Dim xmlFileName As String

    xmlFileName = ""
    
    If m_enableLog Then
        
        Set xmlDoc = New MSXML2.DOMDocument
        
        xmlDoc.async = False
        xmlDoc.loadXML xml
        
        Dim dotPos As Integer
        
        dotPos = InStrRev(m_fileName, ".", -1, vbTextCompare)
        
        If dotPos > 0 Then
            xmlFileName = Left$(m_fileName, dotPos) & CStr(m_xmlFileIndex) & ".xml"
        Else
            xmlFileName = m_fileName & "." & CStr(m_xmlFileIndex) & ".xml"
        End If
        
        m_xmlFileIndex = m_xmlFileIndex + 1
        
        xmlDoc.save xmlFileName
                        
    End If
    
    WriteXml = xmlFileName
    
exitLabel:
    Set xmlDoc = Nothing
    Exit Function

errorLabel:
    ReportError Err, "Log.WriteXml", "Writing to file " & xmlFileName
    GoTo exitLabel
    
End Function

Private Function WriteAbout() As Boolean
    WriteLine "Product: " & App.ProductName
    WriteLine "Version: " & App.Major & "." & App.Minor & ".0." & App.Revision
    WriteLine "Location: " & App.Path & "\" & App.EXEName & ".dll"
    
    WriteAbout = True
End Function

Private Function ReportError(ByRef pobjErr As ErrObject, ByVal functionName As String, ByVal message As String) As Boolean
    ReportError = False
    
    Dim errorMessage As String
    
    errorMessage = _
        "Error: " & message & vbCrLf & _
        "Error Code: " & pobjErr.Number & vbCrLf & _
        "Error Source: " & pobjErr.Source & vbCrLf & _
        "Error Description: " & pobjErr.Description
    
On Error GoTo errorLabel ' This resets the Err object.
    
    If m_showError Then
        MsgBox errorMessage, vbCritical, "Error"
    End If
    
    ReportError = True
    
exitLabel:
    Exit Function

errorLabel:
    GoTo exitLabel
    
End Function
