VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "KFIPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      KFIPrint.cls
'Copyright:     Copyright © 2006 Vertex Financial Services
'Description:   KFI Printing.
'Dependencies:  None
'Issues:
'
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date     Description
'AS     01/02/06 CORE234 Ensure pdfPrint uses correct printer.
'------------------------------------------------------------------------------------------

Option Explicit

Implements ObjectControl

Private Declare Function pdfPrint Lib "WPPdfSDK01" (ByVal filename As String, ByVal password As String, ByVal licname As String, ByVal lickey As String, ByVal liccode As Long, ByVal Options As String) As Integer
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Private Function GetTemporaryFile() As String
Dim fs As New Scripting.FileSystemObject
Dim L As Long
Dim strTempPathName As String
Dim strWork As String * 256

    L = GetTempPath(256, strWork)
    
    If L = 0 Then
        strTempPathName = App.Path & "\"
    Else
        strTempPathName = Left$(strWork, L)
    End If
    
    GetTemporaryFile = strTempPathName & fs.GetTempName()
    Set fs = Nothing
End Function

Public Function PrintKFI(strDocumentType As String, xmlDocContents As IXMLDOMNode, Optional strCompression As String = "", Optional strPrinterName As String = "", Optional strPW As String, Optional intNumberOfCopies As Integer = 1, Optional intTrayNumber_1 As Integer = 0, Optional intTrayNumber_2 As Integer = 0) As Boolean
Const cFunction = "KFIPrintProcessor.PrintKFI"

Dim arrBytes() As Byte
Dim F As Integer
Dim fs As New Scripting.FileSystemObject
Dim intCopies As Integer
Dim strCompressionType As String
Dim strFileName As String
Dim strWork As String
Dim X As Integer
Dim xmlNode As IXMLDOMNode
        
'PDF specifics
Dim lngLicenceCode As Long
Dim strLicenceName As String
Dim strLicenceKey As String
Dim strOptions As String
Dim strPassword As String
        
'MS Word Specifics
Dim objWordApp As Object

'Compression specifics
Dim objDMSCompression As New dmsCompression1
                
    On Error GoTo PrintKFIError:
    
    Select Case strCompression
        Case "KFI", "ZLIB"
            strCompressionType = "ZLIB"
        Case "DMS", "COMPAPI"
            strCompressionType = "COMPAPI"
        Case Else
            strCompressionType = ""
    End Select
    
    intCopies = intNumberOfCopies
    If intCopies < 1 Then
        intCopies = 1
    End If
    
'Extract the processed data from the RESPONSE block
    Set xmlNode = xmlDocContents
    If xmlNode Is Nothing Then
        Err.Raise 61002, cFunction, "No data provided for printing"
    End If
    xmlNode.dataType = "bin.base64"

'Create a temporary file
    strFileName = GetTemporaryFile()

    F = FreeFile()
    Open strFileName For Binary Access Write As F
    If strDocumentType = "rtf" Then
        If strCompressionType = "" Then
            strWork = xmlNode.nodeTypedValue
        Else
            arrBytes = xmlNode.nodeTypedValue
            objDMSCompression.CompressionMethod = strCompressionType
            strWork = objDMSCompression.SafeArrayDecompressToBSTR(arrBytes)
        End If
        Put F, , strWork
        Close F

'Convert rtf to pdf
        strFileName = ConvertToPDF(strFileName)
        If strFileName = "" Then
            Err.Raise 61004, cFunction, "Failed to convert to pdf"
        End If
    Else
        arrBytes = xmlNode.nodeTypedValue
        If strCompressionType <> "" Then
            objDMSCompression.CompressionMethod = strCompressionType
            arrBytes = objDMSCompression.SafeArrayDecompressToSafeArray(arrBytes)
        End If
'Write binary of MS Word .doc to file
        Put F, , arrBytes
        
        Close F
    End If
    
'    App.LogEvent cFunction & " : print type : " & strDocumentType
    Select Case strDocumentType
        Case "doc"
            Set objWordApp = CreateObject("Word.Application")
            If objWordApp Is Nothing Then
                Err.Raise 61003, cFunction, "Microsoft Word not found"
            End If
            objWordApp.DisplayAlerts = 0 'wdAlertsNone
            objWordApp.Options.SaveNormalPrompt = False
            objWordApp.Documents.Open strFileName, False
            If strPrinterName <> "" Then
                objWordApp.WordBasic.FilePrintSetup Printer:=strPrinterName, DoNotSetAsSysDefault:=1
            End If
            
            If intTrayNumber_1 > 0 Then
                objWordApp.ActiveDocument.PageSetup.FirstPageTray = intTrayNumber_1
                If intTrayNumber_2 > 0 Then
                    objWordApp.ActiveDocument.PageSetup.OtherPagesTray = intTrayNumber_2
                Else
                    objWordApp.ActiveDocument.PageSetup.OtherPagesTray = intTrayNumber_1
                End If
            End If
            
            objWordApp.ActiveDocument.PrintOut Background:=False, Copies:=intCopies
        
        Case "rtf", "pdf"
        
            If strPrinterName <> "" Then
                'AS 01/02/06 CORE234 pdfPrint treats spaces (or commas) in the printer name as delimiters,
                'unless the option is double quoted.
                strOptions = """PRINTER=" & strPrinterName & """"
            End If
            If intCopies > 1 Then
                If strOptions <> "" Then
                    strOptions = strOptions & vbCrLf
                End If
                strOptions = strOptions & "COPIES=" & intCopies
            End If
            
            If intTrayNumber_1 > 0 Then
                If strOptions <> "" Then
                    strOptions = strOptions & vbCrLf
                End If
                strOptions = strOptions & "TRAY=" & intTrayNumber_1
            End If
            
            If intTrayNumber_2 > 0 Then
                If strOptions <> "" Then
                    strOptions = strOptions & vbCrLf
                End If
                strOptions = strOptions & "TRAY2=" & intTrayNumber_2
            End If
            
            strLicenceName = "Marlborough-Stirling"
            strLicenceKey = "GAjAd9yU9dk9GA9kyRe"
            strPassword = strPW
            lngLicenceCode = 5384401
            
            X = pdfPrint(strFileName, strPassword, strLicenceName, strLicenceKey, lngLicenceCode, strOptions)
    
    End Select
    PrintKFI = True
    
PrintKFITidyUp:
    On Error Resume Next

'Close any open files
    Close
    If Not objWordApp Is Nothing Then

'Close any Word documents that are open and release the Word object
        If Not objWordApp.ActiveDocument Is Nothing Then
            objWordApp.ActiveDocument.Close
        End If
        objWordApp.Quit
        Set objWordApp = Nothing
    End If

'Delete the temporary file
    fs.DeleteFile strFileName, True
    Set fs = Nothing
    Set objDMSCompression = Nothing
    On Error GoTo 0

    Exit Function

PrintKFIError:
    App.LogEvent cFunction & " : Failed - Error Number : " & Err.Number & " Description : " & Err.Description
    PrintKFI = False
    Resume PrintKFITidyUp:
End Function

Private Function ConvertToPDF(RTFFileName As String) As String
Const cFunction = "KFIPrintProcessor.KFIPrint.ConvertToPDF"

Dim F As Integer
Dim L As Long
Dim objPDF As Object
Dim P As Printer
Dim PDFFileName As String
    
    On Error GoTo ConvertToPDFError:
'This routine converts an input rtf file to a pdf string.
'The routine we use needs a file as input and a file to write the resultant pdf.
'Temporary files are used for this purpose

'Create a temporary file to receive the output
    PDFFileName = GetTemporaryFile()
    
'Obtain a wRTF2PDF object
    On Error Resume Next
    Set objPDF = CreateObject("wPDF_X01.PDFControl")
    Err.Clear
    On Error GoTo ConvertToPDFError:
    If objPDF Is Nothing Then
        App.LogEvent "wPDF_X01.ocx not found", vbLogEventTypeError
        Err.Raise 61006, cFunction, "wPDF_X01.ocx not found"
    End If
    
    objPDF.PDF_Compression = 1 'Deflate
    objPDF.PDF_FontMode = 0
    objPDF.SECURITY_Permission = 0
    objPDF.SECURITY_OwnerPassword = "R3NrstsEs4EdeC7rtff4"
    objPDF.SECURITY_UserPassword = ""
    
    objPDF.INFO_Author = "Marlborough-Stirling"
    objPDF.INFO_Date = Format$(Now(), "dd mmmm yyyy at hh:nn:ss")
                                                                                     
    If objPDF.StartEngine(App.Path & "\wRTF2PDF02.dll", "Marlborough Stirling plc", "R3NrstsEs4EdeC7rtff4", 16575595) = True Then
        L = objPDF.BeginDoc(PDFFileName, 0)
        If L > 0 Then
    
            Debug.Print "RTF2PDF Start " & Timer
            objPDF.ExecIntCommand 1024, 1 'Use printer to determine sizes
            objPDF.ExecIntCommand 1000, 0
            objPDF.ExecStrCommand 1002, RTFFileName
            objPDF.ExecIntCommand 1100, 0
            objPDF.EndDoc
            Debug.Print "RTF2PDF End " & Timer
            ConvertToPDF = PDFFileName
    
        Else
            L = 0
            For Each P In Printers
                L = L + 1
            Next P
            If L = 0 Then
                App.LogEvent "KFIPrint.ConvertToPDF - wRTF2PDF02.BeginDoc requires a printer. There are no known printers on this system", vbLogEventTypeError
            Else
                App.LogEvent "KFIPrint.ConvertToPDF - wRTF2PDF02.BeginDoc failed with return value : 0", vbLogEventTypeError
            End If
        End If
        objPDF.StopEngine
    Else
        Err.Raise 61005, cFunction, "Unable to load the PDF Engine DLL"
    End If
    
ConvertToPDFError:
    Set objPDF = Nothing
    Set P = Nothing
End Function




Private Sub ObjectControl_Activate()

End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = False
End Function


Private Sub ObjectControl_Deactivate()

End Sub


