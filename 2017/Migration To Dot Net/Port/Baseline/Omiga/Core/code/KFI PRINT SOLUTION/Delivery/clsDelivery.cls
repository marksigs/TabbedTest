VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "Delivery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Workfile:      clsDelivery.cls
'Copyright:     Copyright © 2005 Marlborough Stirling
'Description:
'       Handles new print template data extraction.
'
'Dependencies:
'Issues:        Instancing:         MultiUse
'               MTSTransactionMode: UsesTransaction
'               MUST run single-threaded as wRTF2PDF is a single threded dll
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date            Description
'TW     12/07/2005      CORE165     Modified to handle 'fold line'
'------------------------------------------------------------------------------------------

'Output Properties
Dim lngErrorState As Long
Dim strErrorResponse As String

'Input Properties
Dim intPDFFontMode As Integer
Dim intPDFSecurity As Integer
Dim strPDFUserPassword As String
Dim strPDFOwnerPassword As String

Implements ObjectControl

Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Public Function DeliverAsPDF(strRTF As String) As Variant
Const cFunction = "DeliveryProcessor.Delivery.DeliverAsPDF"

Dim arrBytes() As Byte
Dim F As Integer
Dim fs As New Scripting.FileSystemObject
Dim L As Long
Dim objPDF As Object
Dim P As Printer
Dim PDFFileName As String
Dim RTFFileName As String
Dim strTempPathName As String
Dim strWork As String * 256

Dim intFold As Integer '*************
    
    On Error GoTo DeliverAsPDFError:
'This routine converts an input rtf string to a pdf string.
'The routine we use needs a file as input and a file to write the resultant pdf.
'Temporary files are used for this purpose

'Create Temporary Files
    
    L = GetTempPath(256, strWork)
    
    If L = 0 Then
        strTempPathName = App.Path & "\"
    Else
        strTempPathName = Left$(strWork, L)
    End If
    
    RTFFileName = strTempPathName & fs.GetTempName
    PDFFileName = strTempPathName & fs.GetTempName
    
    F = FreeFile()
    Open RTFFileName For Binary Access Write As F
    Put F, , strRTF
    Close F
'TW 12/07/2005 CORE165
'See if there is a fold line.
'This is indicated by a '\shptop' rtf 'control sequence
'followed by a positive number > 0 defining the TWIP location
'for the line. Only one fold line can be handled
    L = 1
    Do
        L = InStr(L, strRTF, "\shptop")
        If L = 0 Then
            Exit Do
        End If
        L = L + 7
        Select Case Mid$(strRTF, L, 1)
            Case "0", "-"
            Case Else
                intFold = Val(Mid$(strRTF, L, 5))
                Exit Do
        End Select
    Loop
'TW 12/07/2005 CORE165 End
    
'Obtain a wRTF2PDF object
    On Error Resume Next
    Set objPDF = CreateObject("wPDF_X01.PDFControl")
    Err.Clear
    On Error GoTo DeliverAsPDFError:
    If objPDF Is Nothing Then
        App.LogEvent "wPDF_X01.ocx not found", vbLogEventTypeError
        Err.Raise 65003, cFunction, "wPDF_X01.ocx not found"
    End If
    
    objPDF.PDF_Compression = 1 'Deflate
''    objPDF.PDF_FontMode = 2 'Embed symbolTrueType fonts
'    objPDF.PDF_FontMode = 0 'Use TrueType fonts - not embedded (default)
'
''    objPDF.SECURITY_Permission = 1 'View only allowed
'    objPDF.SECURITY_Permission = 3 'Printing allowed - no other access (default)
    
    objPDF.PDF_FontMode = intPDFFontMode
    objPDF.SECURITY_Permission = intPDFSecurity
    objPDF.SECURITY_OwnerPassword = strPDFOwnerPassword
    objPDF.SECURITY_UserPassword = strPDFUserPassword
    
    objPDF.INFO_Author = "Marlborough-Stirling"
    objPDF.INFO_Date = Format$(Now(), "dd mmmm yyyy at hh:nn:ss")
                                                                                     
    If objPDF.StartEngine(App.Path & "\wRTF2PDF02.dll", "Marlborough Stirling plc", "R3NrstsEs4EdeC7rtff4", 16575595) = True Then
        L = objPDF.BeginDoc(PDFFileName, 0)
        If L > 0 Then
    
            Debug.Print "RTF2PDF Start " & Timer
            objPDF.ExecIntCommand 1024, 1 'Use printer to determine sizes
            objPDF.ExecIntCommand 1000, 0
            objPDF.ExecStrCommand 1002, RTFFileName
'TW 12/07/2005 CORE165
'Fold line request. (<0=auto, 0=off, >0 = y in twips)
            objPDF.ExecIntCommand 1050, intFold
'TW 12/07/2005 CORE165 End
            objPDF.ExecIntCommand 1100, 0
            objPDF.EndDoc
            Debug.Print "RTF2PDF End " & Timer
    
        Else
            L = 0
            For Each P In Printers
                L = L + 1
            Next P
            If L = 0 Then
                App.LogEvent "Delivery.DeliverAsPDF - wRTF2PDF02.BeginDoc requires a printer. There are no known printers on this system", vbLogEventTypeError
            Else
                App.LogEvent "Delivery.DeliverAsPDF - wRTF2PDF02.BeginDoc failed with return value : 0", vbLogEventTypeError
            End If
        End If
        objPDF.StopEngine
    Else
        Err.Raise 65004, cFunction, "Unable to load the PDF Engine DLL"
    End If
    
    If fs.FileExists(PDFFileName) Then
        
        Set objPDF = Nothing
        L = FileLen(PDFFileName)
        ReDim arrBytes(L - 1)
    
        Open PDFFileName For Binary Access Read As F
        Get F, , arrBytes
        Close F
    
        DeliverAsPDF = arrBytes
    Else
        App.LogEvent "Delivery.DeliverAsPDF - PDF file not found", vbLogEventTypeError
    End If
DeliverAsPDFError:
    
    lngErrorState = Err.Number
    strErrorResponse = Err.Description
'Delete Temporary Files

    If fs.FileExists(RTFFileName) Then
        Kill RTFFileName
    End If
    If fs.FileExists(PDFFileName) Then
        Kill PDFFileName
    End If

    Set objPDF = Nothing
    Set P = Nothing
    Set fs = Nothing
End Function

Public Property Let OwnerPassword(strOwnerPassword As String)
    strPDFOwnerPassword = strOwnerPassword
End Property

Public Property Let UserPassword(strUserPassword As String)
    strPDFUserPassword = strUserPassword
End Property


Public Property Let FontMode(intFontMode As Integer)
    intPDFFontMode = intFontMode
End Property

Public Property Let SecurityLevel(intSecurity As Integer)
    intPDFSecurity = intSecurity
End Property


Public Property Get ErrorResponse() As String
    ErrorResponse = strErrorResponse
End Property

Public Property Get ErrorState() As Long
    ErrorState = lngErrorState
End Property




Private Sub Class_Initialize()
    intPDFFontMode = 0 'Use TrueType fonts - not embedded (default)
    intPDFSecurity = 3 'Printing allowed - no other access (default)
    strPDFUserPassword = ""
    strPDFOwnerPassword = "R3NrstsEs4EdeC7rtff4"
End Sub

Private Sub ObjectControl_Activate()

End Sub


Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function


Private Sub ObjectControl_Deactivate()

End Sub


