VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "axwordclass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Workfile:      axwordclass.cls
'Copyright:     Copyright © 2002 Marlborough Stirling

'Description:
'Dependencies:
'Issues:        Under W2K encountered problems creating the CAB file - Error code 429
'               Check out KB article Q255726 for solution.
'
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date     Description
'DR     07/03/02 Created
'DJB    08/05/02 Many issues resolved including Word activationm many enhancements and bug fixes.
'LD     15/10/02 Word 97 compatible
'DJB,LD 14/02/03 Changed the temporary filename to be a GUID to ensure no file name collision.
'DJB    21/05/03 Fix to change the browser control type to page browse, this corrects bug OPC0006798
'DJB    16/17/03 Change to default window size to the size of the screen.
'TW     16/06/04 Modified XML Property Let to handle pdf files
'AS     13/10/05 CORE202 Added modeless option to DisplayDocumentNative.
'AS     25/04/06 CORE266 Axword: add support for viewing AFP files
'------------------------------------------------------------------------------------------


Option Explicit

Implements IObjectSafety
#If ASYNC_PRINTING Then
'Implements IMessageQueueComponentVC1    ' for message queue listening
#End If

Public Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

'Used in the error reporting to tell us where the error has occured.
Private Const gstrObjectName = "AxwordClass.cls"

Private Const ERROR_PERMISSION_DENIED = 70

Private Sub Class_Initialize()
    On Error Resume Next
       
    ' Initialise variables.
    gintHeight = 0
    gintWidth = 0
    gintCurrentHeight = 480
    gintCurrentWidth = 640
    gintCurrentTop = 0
    gintCurrentLeft = 0
    gFileSaved = False
    gblnError = False
    gblnEdit = False
    gblnHTMLGenerated = False
    gblnInError = False
    gblnWindowShown = False
    gblnResizeableFrame = False
    gblnPersistState = False
    gintFileRetries = 120
    gblnViewAsWord = False
    gblnViewAsPDF = False
    gblnReadOnly = True
    gblnSpellCheckOnSave = True
    gblnSpellCheckWhileEditing = True
    gintPageFit = wdPageFitBestFit
    gblnShowFindFreeText = True
    gblnShowCommandBars = False
    gblnCommandBarsOn = False
    gblnShowPrint = False
    gblnShowPrintDialog = False
    gblnShowProgressBar = False
    gblnShowTrackedChanges = False
    gblnDocumentEdited = False
    gblnDocumentPrinted = False
    gblnDisablePrintOut = False
    gblnModeless = False
    gRequestData.strDocumentID = ""
    gRequestData.strDocumentTitle = ""
    gRequestData.strPrinter = ""
    gRequestData.nCopies = 1
    gRequestData.nFirstPagePrinterTray = wdPrinterDefaultBin
    gRequestData.nOtherPagesPrinterTray = wdPrinterDefaultBin
    gRequestData.bUseDifferentTrayForOtherPages = False

    ClearLastErr

End Sub

Private Sub Class_Terminate()
    ' Ensure word application is closed.
    If (Not gobjWordApplication Is Nothing) Then
        ' Only close if we have no other processes attached to us. (MS bug Q188546)
        If (gobjWordApplication.Documents.Count = 0) Then
            Close_Word_App gobjWordApplication, gstrObjectName
        End If
    End If
End Sub

Private Sub IObjectSafety_GetInterfaceSafetyOptions(ByVal riid As _
    Long, pdwSupportedOptions As Long, pdwEnabledOptions As Long)
    
    On Error Resume Next
    
    Const strFunctionName As String = "IObjectSafety_GetInterfaceSafetyOptions"
    Dim Rc      As Long
    Dim rClsId  As udtGUID
    Dim IID     As String
    Dim bIID()  As Byte

    pdwSupportedOptions = INTERFACESAFE_FOR_UNTRUSTED_CALLER Or _
                          INTERFACESAFE_FOR_UNTRUSTED_DATA

    If (riid <> 0) Then
        CopyMemory rClsId, ByVal riid, Len(rClsId)

        bIID = String$(MAX_GUIDLEN, 0)
        Rc = StringFromGUID2(rClsId, VarPtr(bIID(0)), MAX_GUIDLEN)
        Rc = InStr(1, bIID, vbNullChar) - 1
        IID = Left$(UCase(bIID), Rc)

        Select Case IID
            Case IID_IDispatch, IID_IDispatchEx
                pdwEnabledOptions = IIf(m_fSafeForScripting, INTERFACESAFE_FOR_UNTRUSTED_CALLER, 0)
                Exit Sub
            Case IID_IPersistStorage, IID_IPersistStream, IID_IPersistPropertyBag
                pdwEnabledOptions = IIf(m_fSafeForInitializing, INTERFACESAFE_FOR_UNTRUSTED_DATA, 0)
                Exit Sub
            Case Else
                Err.Raise E_NOINTERFACE
                Exit Sub
        End Select
    End If
    
End Sub

Private Sub IObjectSafety_SetInterfaceSafetyOptions(ByVal riid As _
    Long, ByVal dwOptionsSetMask As Long, ByVal dwEnabledOptions As Long)
    
    On Error Resume Next
    
    Const strFunctionName As String = "IObjectSafety_SetInterfaceSafetyOptions"
    Dim Rc          As Long
    Dim rClsId      As udtGUID
    Dim IID         As String
    Dim bIID()      As Byte

    If (riid <> 0) Then
        CopyMemory rClsId, ByVal riid, Len(rClsId)

        bIID = String$(MAX_GUIDLEN, 0)
        Rc = StringFromGUID2(rClsId, VarPtr(bIID(0)), MAX_GUIDLEN)
        Rc = InStr(1, bIID, vbNullChar) - 1
        IID = Left$(UCase(bIID), Rc)

        Select Case IID
            Case IID_IDispatch, IID_IDispatchEx
                If ((dwEnabledOptions And dwOptionsSetMask) <> INTERFACESAFE_FOR_UNTRUSTED_CALLER) Then
                    Err.Raise E_FAIL
                    Exit Sub
                Else
                    If Not m_fSafeForScripting Then
                        Err.Raise E_FAIL
                    End If
                    Exit Sub
                End If

            Case IID_IPersistStorage, IID_IPersistStream, IID_IPersistPropertyBag
                If ((dwEnabledOptions And dwOptionsSetMask) <> INTERFACESAFE_FOR_UNTRUSTED_DATA) Then
                    Err.Raise E_FAIL
                    Exit Sub
                Else
                    If Not m_fSafeForInitializing Then
                        Err.Raise E_FAIL
                    End If
                    Exit Sub
                End If

            Case Else
                Err.Raise E_NOINTERFACE
                Exit Sub
        End Select
    End If
        
End Sub


Public Property Get xml() As String
    Err.Raise STG_E_UNIMPLEMENTEDFUNCTION, gstrObjectName & "::Get XML"
End Property

Public Property Let xml(ByVal xmlStr As String)
    Err.Raise STG_E_UNIMPLEMENTEDFUNCTION, gstrObjectName & "::Let XML"
End Property

#If MIN_BUILD Then
Public Function DisplayDocument(ByVal xmlStr As String) As String
    Err.Raise STG_E_UNIMPLEMENTEDFUNCTION, gstrObjectName & "::DisplayDocument"
End Function
#Else
Public Function DisplayDocument(ByVal xmlStr As String) As String
    On Error GoTo ExitPoint

    Const strFunctionName = "DisplayDocument"

    DisplayDocument = ""
    
    'All used to save a word file to the users file system so that frmaxord.frm can
    'load it up as an ole object.
    
    If GetRequestData(xmlStr) Then
        Dim strTempPath As String * 512
        
        If GetTempPath(512, strTempPath) > 0 Then
            'Setup the form for viewing, passing through the file name (no extension).
            frmaxword.gstrFile = Replace(strTempPath, Chr(0), "") & GenerateFilename(gRequestData.strDocumentTitle)
            
            ' Initialise globals.
            gblnShowRevisions = False
            
            ' Setup startup status.
            If gblnReadOnly Or gblnViewAsPDF Then
                frmaxword.mstrStartupStatus = "Ready"
            Else
                frmaxword.mstrStartupStatus = "Editing document"
            End If
                                             
    'TW 16/6/2004
            If gblnViewAsPDF Then
                'When using PDFs, other properties which are only relevant to Word must be disabled.
                gblnReadOnly = True
                gblnViewAsWord = False
                gblnSpellCheckOnSave = False
                gblnSpellCheckWhileEditing = False
                gblnShowFindFreeText = False
                gblnShowCommandBars = False
                gblnCommandBarsOn = False
                gblnPDFGenerated = saveBinBase64ToFile(gRequestData.FileContents, gRequestData.nDeliveryType, frmaxword.gstrFile & ".pdf")
            Else
                'Need to save the new XML to file
                saveBinBase64AsWordDoc gRequestData.FileContents, gRequestData.nDeliveryType, frmaxword.gstrFile, gblnViewAsWord
            End If
    'End TW 16/6/2004
            
            ' Show the Gui.
            ' AS 02/02/04 See MSDN Q176468
            ' Modeless forms in an in-process ActiveX DLL are not possible from
            ' Internet Explorer, but have the advantage of showing an icon in the Windows TaskBar.
            ' Therefore, use modeless where possible (not IE).
             ' AS 20/09/04 BMIDS850 Disable Minimize button (must be disabled at design time).
             ' It does not make sense to allowing minimising of the Axword window, because
             ' without a TaskBar icon, it is difficult to locate the window again so it
             ' can be restored/maximised.
            If App.NonModalAllowed Then
                ' Not called from Internet Explorer. TaskBar icon.
                ' AS 01/11/04 change to always be modal. When modeless DisplayDocument returns
                ' immediately, and does not wait for the user to edit the document and exit
                ' from the form, thus DisplayDocument returns the existing document, not the edited
                ' document.
                ' frmaxword.Show vbModeless
                frmaxword.Show vbModal
            Else
                ' Called from Internet Explorer. No TaskBar icon.
                frmaxword.Show vbModal
            End If
        Else
            ' No XML passed.
            Handle_Error Err, gstrObjectName, strFunctionName, "No XML document file passed"
        End If
    End If
    
    If ConvertBase64ToBase64Compressed(gstrObjectName, gRequestData.FileContents) Then
        DisplayDocument = gRequestData.FileContents.strBinBase64
    End If
           
ExitPoint:
    Handle_Error Err, gstrObjectName, strFunctionName
    
End Function
#End If

Public Function DisplayDocumentNative(ByVal xmlStr As String) As String
    On Error GoTo ExitPoint
    Dim bSuccess As Boolean
    Dim strFileName As String
    
    Const strFunctionName = "DisplayDocumentNative"

    DisplayDocumentNative = ""
    
    'All used to save a word file to the users file system so that frmaxord.frm can
    'load it up as an ole object.
    
    bSuccess = False
    If GetRequestData(xmlStr) Then
        Dim strTempPath As String * 512
        
        If GetTempPath(512, strTempPath) > 0 Then
            strFileName = Replace(strTempPath, Chr(0), "") & GenerateFilename(gRequestData.strDocumentTitle) & gRequestData.strFileExtension
            
            If saveBinBase64ToFile(gRequestData.FileContents, gRequestData.nDeliveryType, strFileName) Then
                
                If gblnModeless Then
                    ' If mode less then we do not wait for the user to unlock the file, and
                    ' therefore we can not pick up any edits to the file, thus make the file
                    ' read only to prevent the user editing it.
                    gblnReadOnly = True
                End If
                
                If gblnReadOnly Then
                    SetAttr strFileName, vbReadOnly
                End If
                
                Dim varOldFileDateTime As Variant
                Dim varNewFileDateTime As Variant
                varOldFileDateTime = FileDateTime(strFileName)

                Dim lProcessId As Long
                lProcessId = ShellExecute(strFileName)
                If lProcessId > 0 Then
                    ' If mode less then do not wait for launched application to unlock the file,
                    ' but return immediately to the caller.
                    If Not gblnModeless Then
                        Dim arrFile() As Byte
                        
                        arrFile = ReadFileOnProcessExit(strFileName, lProcessId)
                        
                        varNewFileDateTime = FileDateTime(strFileName)
                        
                        If varOldFileDateTime <> varNewFileDateTime Then
                            gblnDocumentEdited = True
                        End If
                        
                        If ConvertBinToBase64(gstrObjectName, arrFile, gRequestData.FileContents, IIf(Len(gRequestData.FileContents.strCompressionMethod) > 0, True, False)) Then
                            DisplayDocumentNative = gRequestData.FileContents.strBinBase64
                        End If
                    Else
                        DisplayDocumentNative = gRequestData.FileContents.strBinBase64
                    End If
                Else
                    Err.Raise vbObjectError, gstrObjectName & "::" & strFunctionName, "Unable to execute application associated with " & strFileName
                End If
            Else
                Err.Raise vbObjectError, gstrObjectName & "::" & strFunctionName, "Unable to display " & strFileName
            End If
        Else
            ' No XML passed.
            Err.Raise vbObjectError, gstrObjectName & "::" & strFunctionName, "No XML document file passed"
        End If
    End If
    
ExitPoint:
    
    ' Delete any temporary file.
    Call SafeDeleteFile(strFileName)
    
    Handle_Error Err, gstrObjectName, strFunctionName
        
End Function

Private Function ReadFileOnProcessExit(ByVal strFileName As String, ByVal lProcessId As Long) As Byte()

On Error GoTo ErrorHandler

    Dim arrFile() As Byte
    Dim bLoop As Boolean
    Dim bWaitingForProcessToOpenFile As Boolean
    Dim nIteration As Integer
    
    Sleep 1000   ' Give a bit of time for the process to open the file.
    
    bWaitingForProcessToOpenFile = True
    nIteration = 0
    bLoop = True
    While bLoop
        nIteration = nIteration + 1
        Dim hwnd As Long
        Dim hFile As Integer
        Dim bLockFile As Boolean
               
        hwnd = FindProcessWindow(lProcessId)
        
        ' Word (.doc and .rtf) and Acrobat (.pdf) obey file locking, so we can tell when the file is
        ' no longer being viewed (in Word or Acrobat) by using file locks.
        bLockFile = True
        If gRequestData.nDeliveryType = DELIVERYTYPE_UNK And hwnd <> 0 Then
            ' For other applications, e.g., Internet Explorer (.afp) check for the presence of the process
            ' window to tell if the file is still being viewed. Note: this method does not work with Word,
            ' as FindProcessWindow returns 0. Although it does work with Acrobat, the file locking method
            ' is used instead for historical reasons.
            bLockFile = False
        End If

        hFile = FreeFile()
       
        Dim bIsFileOpen As Boolean
        bIsFileOpen = False
        
        If bLockFile Then
            On Error Resume Next
            Open strFileName For Binary Access Read Lock Read Write As hFile
            If Err.Number = ERROR_PERMISSION_DENIED Then
                bIsFileOpen = True
            ElseIf Err.Number = 0 Then
                bIsFileOpen = False
            Else
                GoTo ErrorHandler
            End If
            On Error GoTo ErrorHandler
        Else
            If hwnd <> 0 Then
                bIsFileOpen = True
            Else
                bIsFileOpen = False
            End If
        End If
        
        If bIsFileOpen Then
            ' Process has opened the file.
            If bWaitingForProcessToOpenFile Then
                ' First time, so activate application window, just in case it is not already active.
                ' This may error, e.g., Microsoft Word.
                On Error Resume Next
                Call AppActivate(lProcessId)
                On Error GoTo ErrorHandler
            End If
            bWaitingForProcessToOpenFile = False
            Sleep 500
        Else
            If bWaitingForProcessToOpenFile And nIteration < 120 Then
                ' Process has not opened file yet, so continue waiting.
                If bLockFile Then
                    ' File is not locked because Word/Acrobat has not opened it yet.
                    Close hFile
                End If
                Sleep 500
                
            Else
                ' Process has finished with the file, so read it in.
                
                Dim lFileLength As Long
                lFileLength = FileLen(strFileName)
                If Not bLockFile Then
                    Open strFileName For Binary Access Read Lock Read Write As hFile
                End If
                If gRequestData.nDeliveryType = DELIVERYTYPE_RTF Then
                    ' RTF must be read in as text, not binary.
                    Dim strFile As String
                    strFile = Space(lFileLength)
                    Get hFile, , strFile
                    arrFile = strFile
                Else
                    ReDim arrFile(lFileLength - 1)
                    Get hFile, , arrFile
                End If
                               
                Close hFile

                ReadFileOnProcessExit = arrFile

                bLoop = False
            End If
        End If
        DoEvents
    Wend
    
    Exit Function
    
ErrorHandler:

    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext

End Function

' AS 27/08/04 Background print of document. Does not invoke frmAxword.
Public Function PrintDocument(ByVal xmlStr As String) As Boolean
    On Error GoTo ExitPoint
    
    g_Log.OpenLog
    
    g_Log.WriteLineXml "->PrintDocument", xmlStr
    
    Dim bSuccess As Boolean
    bSuccess = False
    
    Const strFunctionName As String = "PrintDocument"
    
    If GetRequestData(xmlStr) Then
        bSuccess = PrintDocumentRequestData(gRequestData)
    End If
       
ExitPoint:
    
    g_Log.WriteLine "<-PrintDocument = " & IIf(bSuccess, "True", "False")
    
    g_Log.CloseLog
            
    PrintDocument = bSuccess
End Function

Private Function PrintDocumentRequestData(ByRef RequestData As REQUEST_DATA) As Boolean
    On Error GoTo ExitPoint
    
    g_Log.WriteLine "->PrintDocumentRequestData"
    
    Dim strFileName As String
    Dim bSuccess As Boolean
    bSuccess = False
    
    Const strFunctionName As String = "PrintDocumentRequestData"
    
    Dim strTempPath As String * 512
    
    If GetTempPath(512, strTempPath) > 0 Then
        strFileName = Replace(strTempPath, Chr(0), "") & GenerateFilename(gRequestData.strDocumentTitle) & RequestData.strFileExtension
        
        If saveBinBase64ToFile(RequestData.FileContents, RequestData.nDeliveryType, strFileName) Then
            If RequestData.nDeliveryType = DELIVERYTYPE_PDF Then
                bSuccess = PrintPDFDocument(Nothing, strFileName, RequestData.strDocumentTitle, False, gblnShowPrintDialog, gblnShowProgressBar, RequestData.strPrinter, RequestData.nFirstPagePrinterTray, RequestData.nOtherPagesPrinterTray, RequestData.bUseDifferentTrayForOtherPages, RequestData.nCopies)
            ElseIf RequestData.nDeliveryType = DELIVERYTYPE_RTF Then
                bSuccess = PrintPDFDocument(Nothing, strFileName, RequestData.strDocumentTitle, True, gblnShowPrintDialog, gblnShowProgressBar, RequestData.strPrinter, RequestData.nFirstPagePrinterTray, RequestData.nOtherPagesPrinterTray, RequestData.bUseDifferentTrayForOtherPages, RequestData.nCopies)
            Else
                bSuccess = PrintWordDocument(Nothing, Nothing, strFileName, RequestData.strDocumentTitle, gblnShowPrintDialog, gblnShowProgressBar, RequestData.strPrinter, RequestData.nFirstPagePrinterTray, RequestData.nOtherPagesPrinterTray, RequestData.bUseDifferentTrayForOtherPages, RequestData.nCopies)
            End If
        End If
    End If
       
ExitPoint:
    
    ' Delete any temporary file.
    Call SafeDeleteFile(strFileName)
    
    Handle_Error Err, gstrObjectName, strFunctionName
        
    g_Log.WriteLine "<-PrintDocumentRequestData = " & CStr(bSuccess)
    
    PrintDocumentRequestData = bSuccess
End Function

Public Function GetPrintersAsXML() As String
    On Error GoTo ExitPoint
    Const strFunctionName As String = "PrintDocument"

    Dim bSuccess As Boolean
    Dim PrinterData() As PRINTER_DATA

    bSuccess = GetPrinters(PrinterData)

    If bSuccess Then
        Dim xmlDoc As MSXML2.DOMDocument
        Set xmlDoc = New MSXML2.DOMDocument

        If xmlDoc Is Nothing Then
            Err.Raise vbObjectError, gstrObjectName & "::" & strFunctionName, "Failed to create XML Parser - check MSXML is installed"
        End If

        xmlDoc.async = False

        Dim nodeResponse As IXMLDOMElement
        Set nodeResponse = xmlDoc.createNode(NODE_ELEMENT, "RESPONSE", "")
        xmlDoc.appendChild nodeResponse

        Dim nodePrinters As IXMLDOMElement
        Set nodePrinters = xmlDoc.createNode(NODE_ELEMENT, "PRINTERS", "")
        nodeResponse.appendChild nodePrinters

        Dim nPrinter As Integer
        For nPrinter = 0 To UBound(PrinterData)
            On Error Resume Next
            
            Dim nodePrinter As IXMLDOMElement
            Set nodePrinter = xmlDoc.createNode(NODE_ELEMENT, "PRINTER", "")

            nodePrinter.setAttribute "DEVICENAME", PrinterData(nPrinter).strDeviceName
            nodePrinter.setAttribute "DEFAULT", IIf(PrinterData(nPrinter).bDefault, "TRUE", "FALSE")
            nodePrinter.setAttribute "BINS", PrinterData(nPrinter).nBins
            nodePrinter.setAttribute "DEFAULTBIN", PrinterData(nPrinter).nDefaultBin

            nodePrinter.setAttribute "DEFAULTBIN", PrinterData(nPrinter).nDefaultBin
            nodePrinter.setAttribute "DEFAULTBIN", PrinterData(nPrinter).nDefaultBin

            If PrinterData(nPrinter).nBins > 0 Then
                Dim nBin As Integer
                For nBin = 0 To UBound(PrinterData(nPrinter).nBinNumbers)
                    Dim nodeBin As IXMLDOMElement
                    Set nodeBin = xmlDoc.createNode(NODE_ELEMENT, "BIN", "")
    
                    nodeBin.setAttribute "BINNUMBER", PrinterData(nPrinter).nBinNumbers(nBin)
                    nodeBin.setAttribute "BINNAME", PrinterData(nPrinter).strBinNames(nBin)
    
                    nodePrinter.appendChild nodeBin
                Next
            End If
            
            nodePrinters.appendChild nodePrinter
        Next

        On Error GoTo ExitPoint
        
        GetPrintersAsXML = xmlDoc.xml

    End If

ExitPoint:

    Set nodeBin = Nothing
    Set nodePrinter = Nothing
    Set nodePrinters = Nothing
    Set xmlDoc = Nothing

    Handle_Error Err, gstrObjectName, strFunctionName
End Function

Public Function ConvertFileToBase64(ByVal strFileName As String, Optional ByVal strCompressionMethod As String = "", Optional ByVal blnDeleteFile As Boolean = False) As String
On Error GoTo ExitPoint
    Const strFunctionName As String = "ConvertFileToBase64"
           
    Dim arrData() As Byte
    Dim lFileLength As Long
    lFileLength = FileLen(strFileName)
  
    Dim hFile As Integer
    hFile = FreeFile()
    Open strFileName For Binary Access Read As hFile
    If LCase$(Right(strFileName, 4)) = ".rtf" And strCompressionMethod = "" Then
        Dim strFile As String
        strFile = Space(lFileLength)
        Get hFile, , strFile
        arrData = strFile
    Else
        ReDim arrData(lFileLength - 1)
        Get hFile, , arrData
    End If
    Close hFile
        
    Dim FileContents As FILE_CONTENTS
    FileContents.strCompressionMethod = strCompressionMethod
    Call ConvertBinToBase64(gstrObjectName, arrData, FileContents, bCompress:=IIf(Len(strCompressionMethod) > 0, True, False))
    
    If blnDeleteFile Then
        Kill strFileName
    End If
    
    ConvertFileToBase64 = FileContents.strBinBase64
   
ExitPoint:
    
    Handle_Error Err, gstrObjectName, strFunctionName
    
End Function

Public Function ConvertBase64ToFile(ByVal strBinBase64, ByVal strFileName As String, Optional ByVal strCompressionMethod As String = "") As Boolean
On Error GoTo ExitPoint
    Const strFunctionName As String = "ConvertFileToBase64"
    
    Dim bSuccess As Boolean
    bSuccess = False
    
    Dim arrData() As Byte
    Dim FileContents As FILE_CONTENTS
    FileContents.strBinBase64 = strBinBase64
    FileContents.strCompressionMethod = strCompressionMethod
    FileContents.bCompressed = IIf(Len(strCompressionMethod) > 0, True, False)
    FileContents.bCompressedArray = False
    arrData = ConvertBase64ToBin(gstrObjectName, FileContents, bDecompress:=FileContents.bCompressed)
        
    On Error Resume Next
    Kill strFileName
    On Error GoTo ExitPoint
    
    Dim hFile As Integer
    hFile = FreeFile()
    Open strFileName For Binary Access Write Lock Read Write As hFile
    Put hFile, , arrData
    Close hFile
        
    bSuccess = True
    
ExitPoint:
    
    Handle_Error Err, gstrObjectName, strFunctionName
    
    ConvertBase64ToFile = bSuccess
End Function

Private Function GetRequestData(ByVal xmlStr As String) As Boolean
On Error GoTo ExitPoint
    
    g_Log.WriteLine "->GetRequestData"
    
    Dim bSuccess As Boolean
    Dim nRequestCopies As Integer
    
    nRequestCopies = -1
    
    bSuccess = False
    
    Const strFunctionName As String = "GetRequestData"
    
    gRequestData.strOperation = ""
    gRequestData.strDocumentID = ""
    gRequestData.strDocumentTitle = ""
    gRequestData.strPrinter = ""
    gRequestData.nCopies = 1
    gRequestData.nFirstPagePrinterTray = wdPrinterDefaultBin
    gRequestData.nOtherPagesPrinterTray = wdPrinterDefaultBin
    gRequestData.bUseDifferentTrayForOtherPages = False
    gRequestData.nDeliveryType = DELIVERYTYPE_DOC
    gRequestData.strFileExtension = "" ' Must be empty string.
    gRequestData.FileContents.strBinBase64 = ""
    gRequestData.FileContents.strCompressionMethod = ""
    gRequestData.FileContents.bCompressed = False
    gRequestData.FileContents.bCompressedArray = False
        
    If Len(xmlStr) > 0 Then
        Dim xmlDoc As MSXML2.DOMDocument
        Set xmlDoc = New MSXML2.DOMDocument
        
        If xmlDoc Is Nothing Then
            Err.Raise vbObjectError, gstrObjectName & "::" & strFunctionName, "Failed to create XML Parser - check MSXML is installed"
        End If
               
        xmlDoc.async = False
        xmlDoc.loadXML xmlStr
        If xmlDoc.parseError.errorCode <> 0 Then
            Err.Raise _
                vbObjectError, _
                gstrObjectName & "::" & strFunctionName, _
                "Error loading XML. " & vbCrLf & _
                "parseError.errorCode: " & xmlDoc.parseError.errorCode & vbCrLf & _
                "parseError.reason: " & xmlDoc.parseError.reason
        End If
        
        Dim xmlRequestNode As IXMLDOMElement
        Set xmlRequestNode = xmlDoc.documentElement
        
        Dim xmlAttribute As IXMLDOMAttribute
        Set xmlAttribute = xmlRequestNode.Attributes.getNamedItem("OPERATION")
        If Not xmlAttribute Is Nothing Then
            If Len(xmlAttribute.Text) > 0 Then
                gRequestData.strOperation = UCase(xmlAttribute.Text)
            End If
        End If
        
        'Check XML itself - the structure of the document.
        Dim xmlDocContentsNode As IXMLDOMNode
        Set xmlDocContentsNode = xmlDoc.selectSingleNode("REQUEST/DOCUMENTCONTENTS")
        If xmlDocContentsNode Is Nothing Then
            Set xmlDocContentsNode = xmlDoc.selectSingleNode("REQUEST/PRINTDOCUMENTDATA")
            If xmlDocContentsNode Is Nothing Then
                Err.Raise vbObjectError, gstrObjectName & "::" & strFunctionName, "Invalid XML request - missing mandatory node: REQUEST/DOCUMENTCONTENTS"
            End If
        End If
        
        'Check XML itself - attribute that holds the file contents.
        Set xmlAttribute = xmlDocContentsNode.Attributes.getNamedItem("FILECONTENTS")
        If xmlAttribute Is Nothing Then
            Err.Raise vbObjectError, gstrObjectName & "::" & strFunctionName, "Invalid XML request - missing mandatory node: REQUEST/DOCUMENTCONTENTS/FILECONTENTS"
        End If
        'Now get the XML data.
        gRequestData.FileContents.strBinBase64 = xmlAttribute.Text
       
        Dim xmlControlDataNode As IXMLDOMNode
        Set xmlControlDataNode = xmlDoc.selectSingleNode("REQUEST/CONTROLDATA")
        If Not xmlControlDataNode Is Nothing Then
            Set xmlAttribute = xmlControlDataNode.Attributes.getNamedItem("COMPRESSIONMETHOD")
            If Not xmlAttribute Is Nothing Then
                gRequestData.FileContents.strCompressionMethod = xmlAttribute.Text
                If Len(gRequestData.FileContents.strCompressionMethod) > 0 Then
                    gRequestData.FileContents.bCompressed = True
                End If
            End If
             
            Set xmlAttribute = xmlControlDataNode.Attributes.getNamedItem("DOCUMENTID")
            If Not xmlAttribute Is Nothing Then
                If Len(xmlAttribute.Text) > 0 Then
                    gRequestData.strDocumentID = xmlAttribute.Text
                End If
            End If
                       
            Set xmlAttribute = xmlControlDataNode.Attributes.getNamedItem("DOCUMENTTITLE")
            If Not xmlAttribute Is Nothing Then
                If Len(xmlAttribute.Text) > 0 Then
                    gRequestData.strDocumentTitle = xmlAttribute.Text
                End If
            End If

            Set xmlAttribute = xmlControlDataNode.Attributes.getNamedItem("PRINTER")
            If Not xmlAttribute Is Nothing Then
                If Len(xmlAttribute.Text) > 0 Then
                    gRequestData.strPrinter = xmlAttribute.Text
                End If
            End If
                            
            Set xmlAttribute = xmlControlDataNode.Attributes.getNamedItem("FIRSTPAGEPRINTERTRAY")
            If Not xmlAttribute Is Nothing Then
                If Len(xmlAttribute.Text) > 0 Then
                    gRequestData.nFirstPagePrinterTray = CInt(xmlAttribute.Text)
                End If
            End If
            
            Set xmlAttribute = xmlControlDataNode.Attributes.getNamedItem("OTHERPAGESPRINTERTRAY")
            If Not xmlAttribute Is Nothing Then
                If Len(xmlAttribute.Text) > 0 Then
                    gRequestData.nOtherPagesPrinterTray = CInt(xmlAttribute.Text)
                End If
            End If
            
            Set xmlAttribute = xmlControlDataNode.Attributes.getNamedItem("COPIES")
            If Not xmlAttribute Is Nothing Then
                If Len(xmlAttribute.Text) > 0 Then
                    nRequestCopies = CInt(xmlAttribute.Text)
                    gRequestData.nCopies = nRequestCopies
                End If
            End If
            
            Set xmlAttribute = xmlControlDataNode.Attributes.getNamedItem("DELIVERYTYPE")
            If Not xmlAttribute Is Nothing Then
                If Len(xmlAttribute.Text) > 0 Then
                    gRequestData.nDeliveryType = CInt(xmlAttribute.Text)
                End If
            End If
            
            Set xmlAttribute = xmlRequestNode.Attributes.getNamedItem("FILEEXTENSION")
            If Not xmlAttribute Is Nothing Then
                If Len(xmlAttribute.Text) > 0 Then
                    gRequestData.strFileExtension = xmlAttribute.Text
                End If
            End If
            If Len(gRequestData.strFileExtension) = 0 Then
                If gRequestData.nDeliveryType = DELIVERYTYPE_PDF Then
                    gRequestData.strFileExtension = ".pdf"
                ElseIf gRequestData.nDeliveryType = DELIVERYTYPE_RTF Then
                    gRequestData.strFileExtension = ".rtf"
                Else
                    gRequestData.strFileExtension = ".doc"
                End If
            Else
                If LCase(gRequestData.strFileExtension) = ".pdf" Then
                    gRequestData.nDeliveryType = DELIVERYTYPE_PDF
                ElseIf LCase(gRequestData.strFileExtension) = ".rtf" Then
                    gRequestData.nDeliveryType = DELIVERYTYPE_RTF
                ElseIf LCase(gRequestData.strFileExtension) = ".doc" Then
                    gRequestData.nDeliveryType = DELIVERYTYPE_DOC
                Else
                    gRequestData.nDeliveryType = DELIVERYTYPE_UNK
                End If
            End If
            
            If gRequestData.nDeliveryType = DELIVERYTYPE_PDF Then
                gblnViewAsPDF = True
                gblnViewAsWord = False
            End If
                                    
        End If
        
        bSuccess = True
        
    Else
        ' No XML passed.
        Err.Raise vbObjectError, gstrObjectName & "::" & strFunctionName, "No XML request"
    End If
    
    If gblnPersistState Then
        ' Load request defaults from registry for this document. This will override any
        ' equivalent settings in the input request.
        LoadPrintDialogState
    End If
    
    ' AS 12/05/2005 CORE124 Do not override COPIES in request.
    If nRequestCopies > -1 Then
        gRequestData.nCopies = nRequestCopies
    End If
    
ExitPoint:
    Set xmlAttribute = Nothing
    Set xmlDocContentsNode = Nothing
    Set xmlControlDataNode = Nothing
    Set xmlDoc = Nothing
        
    If Err.Number <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
    
    g_Log.WriteLine "<-GetRequestData = " & CStr(bSuccess)
    
    GetRequestData = bSuccess
End Function

Private Function GenerateFilename(Optional ByVal strPrefix As String = "") As String
    ' Generate a GUID filename.
    On Error GoTo ExitPoint
    
    Const strFunctionName As String = "GenerateFilename()"
    
    Dim Guid As String
    Guid = CreateGUID
    If strPrefix <> "" Then
        GenerateFilename = strPrefix & "." & Guid
    Else
        GenerateFilename = Guid
    End If
    
ExitPoint:
    Handle_Error Err, gstrObjectName, strFunctionName
End Function

Public Property Get ReadOnly() As Variant
    ReadOnly = gblnReadOnly
End Property


Public Property Let ReadOnly(ByVal vNewValue As Variant)
    gblnReadOnly = vNewValue
End Property


Public Property Get FileSize() As Variant
    FileSize = gFileSize
End Property


Public Property Get FileSaved() As Variant
    FileSaved = gFileSaved
End Property

Public Property Get ErrorHappened() As Variant
    ErrorHappened = gblnError
End Property

Public Property Get FormWidth() As Variant
    FormWidth = gintWidth
End Property

Public Property Let FormWidth(ByVal vNewValue As Variant)
    ' Check new value if numeric.
    If (IsNumeric(vNewValue)) Then
        ' Set width.
        gintWidth = CInt(vNewValue)
    End If
End Property

Public Property Get FormHeight() As Variant
    FormWidth = gintHeight
End Property

Public Property Let FormHeight(ByVal vNewValue As Variant)
    ' Check new value if numeric.
    If (IsNumeric(vNewValue)) Then
        ' Set width.
        gintHeight = CInt(vNewValue)
    End If
End Property

Public Property Get ResizeableFrame() As Variant
    ResizeableFrame = gblnResizeableFrame
End Property

Public Property Let ResizeableFrame(ByVal vNewValue As Variant)
    gblnResizeableFrame = vNewValue
End Property

Public Property Get PersistState() As Variant
    PersistState = gblnPersistState
End Property

Public Property Let PersistState(ByVal vNewValue As Variant)
    gblnPersistState = vNewValue
End Property

Public Property Get FileRetries() As Variant
    FileRetries = gintFileRetries
End Property

Public Property Let FileRetries(ByVal vNewValue As Variant)
    ' Check new value if numeric.
    If (IsNumeric(vNewValue)) And CInt(vNewValue) > 0 Then
        gintFileRetries = CInt(vNewValue)
    End If
End Property

Public Property Get ViewAsWord() As Variant
    ViewAsWord = gblnViewAsWord
End Property

Public Property Let ViewAsWord(ByVal vNewValue As Variant)
    gblnViewAsWord = vNewValue
End Property

Public Property Get SpellCheckWhileEditing() As Variant
    SpellCheckWhileEditing = gblnSpellCheckWhileEditing
End Property

Public Property Let SpellCheckWhileEditing(ByVal vNewValue As Variant)
    gblnSpellCheckWhileEditing = vNewValue
End Property

Public Property Get SpellCheckOnSave() As Variant
    SpellCheckOnSave = gblnSpellCheckOnSave
End Property

Public Property Let SpellCheckOnSave(ByVal vNewValue As Variant)
    gblnSpellCheckOnSave = vNewValue
End Property

Public Property Get PageFit() As Variant
    PageFit = gintPageFit
End Property

Public Property Let PageFit(ByVal vNewValue As Variant)
    If (IsNumeric(vNewValue)) And CInt(vNewValue) >= wdPageFitNone And CInt(vNewValue) <= wdPageFitBestFit Then
        gintPageFit = CInt(vNewValue)
    End If
End Property

Public Property Get ShowFindFreeText() As Variant
    ShowFindFreeText = gblnShowFindFreeText
End Property

Public Property Let ShowFindFreeText(ByVal vNewValue As Variant)
    gblnShowFindFreeText = vNewValue
End Property

Public Property Get ShowCommandBars() As Variant
    ShowCommandBars = gblnShowCommandBars
End Property

Public Property Let ShowCommandBars(ByVal vNewValue As Variant)
    gblnShowCommandBars = vNewValue
End Property

Public Property Get ViewAsPDF() As Variant
    ViewAsPDF = gblnViewAsPDF
End Property

Public Property Let ViewAsPDF(ByVal vNewValue As Variant)
    gblnViewAsPDF = vNewValue
End Property

Public Property Get ShowPrint() As Variant
    ShowPrint = gblnShowPrint
End Property

Public Property Let ShowPrint(ByVal vNewValue As Variant)
    gblnShowPrint = vNewValue
End Property

Public Property Get ShowPrintDialog() As Variant
    ShowPrintDialog = gblnShowPrintDialog
End Property

Public Property Let ShowPrintDialog(ByVal vNewValue As Variant)
    gblnShowPrintDialog = vNewValue
End Property

Public Property Get ShowProgressBar() As Variant
    ShowProgressBar = gblnShowProgressBar
End Property

Public Property Let ShowProgressBar(ByVal vNewValue As Variant)
    gblnShowProgressBar = vNewValue
End Property

Public Property Get ShowTrackedChanges() As Variant
    ShowTrackedChanges = gblnShowTrackedChanges
End Property

Public Property Let ShowTrackedChanges(ByVal vNewValue As Variant)
    gblnShowTrackedChanges = vNewValue
End Property

Public Property Get DocumentEdited() As Variant
    DocumentEdited = gblnDocumentEdited
End Property

Public Property Get DocumentPrinted() As Variant
    DocumentPrinted = gblnDocumentPrinted
End Property

Public Property Get DisablePrintOut() As Variant
    DisablePrintOut = gblnDisablePrintOut
End Property

Public Property Let DisablePrintOut(ByVal vNewValue As Variant)
    gblnDisablePrintOut = vNewValue
End Property

Public Sub ClearLastErr()
    gLastErrObject.Number = 0
    gLastErrObject.Source = ""
    gLastErrObject.Description = ""
End Sub

Public Function IsLastErr() As Boolean
    IsLastErr = IIf(Not gLastErrObject.Number = 0, True, False)
End Function

Public Property Get LastErrNumber() As Long
    LastErrNumber = gLastErrObject.Number
End Property

Public Property Get LastErrSource() As String
    LastErrSource = gLastErrObject.Source
End Property

Public Property Get LastErrDescription() As String
    LastErrDescription = gLastErrObject.Description
End Property

#If ASYNC_PRINTING Then
Public Function IMessageQueueComponentVC1_OnMessage(ByVal xmlStr As String) As MESSQ_RESP
On Error GoTo ExitPoint
    Dim MQResp As MESSQ_RESP
    MQResp = MESSQ_RESP_RETRY_MOVE_MESSAGE

    If GetRequestData(xmlStr) Then
        Select Case gRequestData.strOperation
        Case "NEWDOCUMENT"
            'FIXAS
        Case "PRINTDOCUMENT"
            gblnShowPrintDialog = False
            If PrintDocumentRequestData(gRequestData) Then
                MQResp = MESSQ_RESP_SUCCESS
            End If
        End Select
    End If
    
ExitPoint:
    IMessageQueueComponentVC1_OnMessage = MQResp
End Function
#End If

Public Property Get EnableLog() As Boolean
    If Not g_Log Is Nothing Then
        EnableLog = g_Log.EnableLog
    Else
        EnableLog = False
    End If
End Property

Public Property Let EnableLog(ByVal enable As Boolean)
    If Not g_Log Is Nothing Then
        g_Log.EnableLog = enable
    End If
End Property

Public Property Get LogFileName() As String
    If Not g_Log Is Nothing Then
        LogFileName = g_Log.FileName
    Else
        LogFileName = ""
    End If
End Property

Public Property Let LogFileName(ByVal fName As String)
    If Not g_Log Is Nothing Then
        g_Log.FileName = fName
    End If
End Property

Public Property Get Modeless() As Variant
    Modeless = gblnModeless
End Property

Public Property Let Modeless(ByVal vNewValue As Variant)
    gblnModeless = vNewValue
End Property


