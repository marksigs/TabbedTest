Attribute VB_Name = "CommonFunctions"
'Workfile:      CommonFunctions.bas
'Copyright:     Copyright © 2002 Marlborough Stirling

'Description:   Common functions for use with Axword.
'Dependencies:  None
'Issues:
'
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date     Description
'DJB    08/05/02 Created.
'DJB    09/08/02 Added new subs ConvertBase64ToBin(), ConvertBinToBase64()
'DJB    18/10/02 Tweaked Handle_Error to unload the form after the first error.
'AS     01/02/06 CORE234 Ensure pdfPrint uses correct printer.
'------------------------------------------------------------------------------------------

Option Explicit

'Used in the error reporting to tell us where the error has occured.
Private Const gstrObjectName = "CommonFunctions.bas"

Private Declare Function DeviceCapabilities Lib "winspool.drv" Alias "DeviceCapabilitiesA" (ByVal lpDeviceName As String, ByVal lpPort As String, ByVal iIndex As Long, lpOutput As Any, lpDevMode As DEVMODE) As Long 'ByVal dev As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal strSystemDirectory As String, ByVal uSize As Long) As Long
Private Declare Function pdfPrint Lib "WPPdfSDK01.dll" (ByVal strFileName As String, ByVal strPassword As String, ByVal strLicenceName As String, ByVal strLicenceKey As String, ByVal lLicenceCode As Long, ByVal strOptions As String) As Long

Private Const DC_BINS = 6
Private Const DC_BINNAMES = 12
Private Const CCHDEVICENAME = 32
Private Const CCHFORMNAME = 32

Private Type DEVMODE
    dmDeviceName As String * CCHDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCHFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
    'win95 only below:
    dmICMMethod As Long
    dmICMIntent As Long
    dmMediaType As Long
    dmDitherType As Long
    dmReserved1 As Long
    dmReserved2 As Long
End Type

Public Type PRINT_DOCUMENT_DATA
    nDeliveryType As Integer
    pobjWordApp As Object
    strFileName As String
    strActivePrinter As String
    nFirstPagePrinterTray As Integer
    nOtherPagesPrinterTray As Integer
    bUseDifferentTrayForOtherPages As Boolean
    nCopies As Integer
    blnConvertFromRTF As Boolean
End Type

Public gPrintDocumentData As PRINT_DOCUMENT_DATA

Public Type PRINTER_DATA
    strDeviceName As String
    bDefault As Boolean
    nBins As Integer
    nDefaultBin As Integer
    nBinNumbers() As Integer
    strBinNames() As String
End Type

Public gPrinterData() As PRINTER_DATA

'-----------------------------------------------------------------
'EnumPrinters START
'-----------------------------------------------------------------

Private Type ACL
    AclRevision As Byte
    Sbz1 As Byte
    AclSize As Integer
    AceCount As Integer
    Sbz2 As Integer
End Type

Private Type SECURITY_DESCRIPTOR
    Revision As Byte
    Sbz1 As Byte
    Control As Long
    Owner As Long
    Group As Long
    Sacl As ACL
    Dacl As ACL
End Type

Private Type PRINTER_INFO_1
    flags As Long
    pDescription As Long
    pName As Long
    pComment As Long
End Type

Private Type PRINTER_INFO_2
    pServerName As Long
    pPrinterName As Long
    pShareName As String
    pPortName As Long
    pDriverName As Long
    pComment As Long
    pLocation As Long
    pDevMode As DEVMODE
    pSepFile As Long
    pPrintProcessor As Long
    pDatatype As Long
    pParameters As Long
    pSecurityDescriptor As SECURITY_DESCRIPTOR
    Attributes As Long
    Priority As Long
    DefaultPriority As Long
    StartTime As Long
    UntilTime As Long
    Status As Long
    cJobs As Long
    AveragePPM As Long
End Type

Private Type PRINTER_INFO_3
    pSecurityDescriptor As SECURITY_DESCRIPTOR
End Type

Private Type PRINTER_INFO_4
    pPrinterName As Long
    pServerName As Long
    Attributes As Long
End Type

Private Type PRINTER_INFO_5
    pPrinterName As Long
    pPortName As Long
    Attributes As Long
    DeviceNotSelectedTimeout As Long
    TransmissionRetryTimeout As Long
End Type

'Level:
'   Windows 95: The value can be 1, 2, or 5.
'   Windows NT/Windows 2000: This value can be 1, 2, 4, or 5.
Private Declare Function EnumPrinters Lib "winspool.drv" Alias "EnumPrintersA" (ByVal flags As Long, ByVal name As String, ByVal Level As Long, pPrinterEnum As Byte, ByVal cbBuf As Long, pcbNeeded As Long, pcReturned As Long) As Long
Private Declare Function lstrlenA Lib "kernel32" (ByVal lpString As Long) As Long

Private Const PRINTER_ENUM_DEFAULT = &H1
Private Const PRINTER_ENUM_LOCAL = &H2
Private Const PRINTER_ENUM_CONNECTIONS = &H4
Private Const PRINTER_ENUM_FAVORITE = &H4
Private Const PRINTER_ENUM_NAME = &H8
Private Const PRINTER_ENUM_REMOTE = &H10
Private Const PRINTER_ENUM_SHARED = &H20
Private Const PRINTER_ENUM_NETWORK = &H40

Private Const PRINTER_ENUM_EXPAND = &H4000
Private Const PRINTER_ENUM_CONTAINER = &H8000

Private Const PRINTER_ENUM_ICONMASK = &HFF0000
Private Const PRINTER_ENUM_ICON1 = &H10000
Private Const PRINTER_ENUM_ICON2 = &H20000
Private Const PRINTER_ENUM_ICON3 = &H40000
Private Const PRINTER_ENUM_ICON4 = &H80000
Private Const PRINTER_ENUM_ICON5 = &H100000
Private Const PRINTER_ENUM_ICON6 = &H200000
Private Const PRINTER_ENUM_ICON7 = &H400000
Private Const PRINTER_ENUM_ICON8 = &H800000

Private Const PRINTER_ATTRIBUTE_QUEUED = &H1
Private Const PRINTER_ATTRIBUTE_DIRECT = &H2
Private Const PRINTER_ATTRIBUTE_DEFAULT = &H4
Private Const PRINTER_ATTRIBUTE_SHARED = &H8
Private Const PRINTER_ATTRIBUTE_NETWORK = &H10
Private Const PRINTER_ATTRIBUTE_HIDDEN = &H20
Private Const PRINTER_ATTRIBUTE_LOCAL = &H40

'-----------------------------------------------------------------
'EnumPrinters END
'-----------------------------------------------------------------


#If Not MIN_BUILD Then
Public Sub ShowFormsDesign(ByVal pstrObjectName As String, ByRef pobjWordDoc As Object, Optional ByVal blnShow As Boolean = True)
    On Error GoTo ExitPoint
      
    'Show/hide forms design bar.
    Const strFunctionName As String = "ShowFormsDesign"
       
    If (pobjWordDoc.CommandBars("Exit Design Mode").Visible And Not blnShow) Or (Not pobjWordDoc.CommandBars("Exit Design Mode").Visible And blnShow) Then
        On Error Resume Next
        pobjWordDoc.ToggleFormsDesign
        'Ignore any error - not important.
        Err.Clear
    End If
        
ExitPoint:
    Handle_Error Err, pstrObjectName, strFunctionName, "", False, False
End Sub

Public Sub ShowCommandBars(ByVal blnShowCommandBars As Boolean, ByVal pstrObjectName As String, ByRef pobjWordDoc As Object)
    On Error GoTo ExitPoint
    
    Const strFunctionName As String = "ShowCommandBars"
    
    'AS 03/02/04 Show all command bars, not just Formatting.
    'Display required command bars.
    'pobjWordDoc.CommandBars("Formatting").Position = msoBarTop
    'pobjWordDoc.CommandBars("Formatting").Protection = msoBarNoChangeDock
    'pobjWordDoc.CommandBars("Formatting").Visible = True
    
    If (blnShowCommandBars And Not gblnCommandBarsOn) Or (Not blnShowCommandBars And gblnCommandBarsOn) Then
        ' This is a toggle option, so call it once to show the
        ' toolbars and once to hide them. This works with Internet Explorer 5
        ' but often fails to work properly with earlier versions.
        If (frmaxword.web1.QueryStatusWB(OLECMDID_HIDETOOLBARS) And OLECMDF_SUPPORTED) = OLECMDF_SUPPORTED Then
            frmaxword.web1.ExecWB OLECMDID_HIDETOOLBARS, OLECMDEXECOPT_DONTPROMPTUSER
        End If
    End If
    
    'AS 05/04/04 Reviewing toolbar introduced in Word 2002 does not obey toggling.
    '10 = Word 2002, 11 = Word 2003
    If glWordVersion >= 10 Then
        Dim objToolBar As Object
        Set objToolBar = pobjWordDoc.CommandBars("Reviewing")
        If Not objToolBar Is Nothing Then
            objToolBar.Visible = gblnShowCommandBars
        End If
    End If
    
    gblnCommandBarsOn = Not gblnCommandBarsOn
   
ExitPoint:
    Handle_Error Err, pstrObjectName, strFunctionName, "", False, False
End Sub
#End If

Public Function PrintWordDocument( _
    ByRef frmParent As Form, _
    ByRef pobjWordApp As Object, _
    ByVal strFileName As String, _
    Optional ByVal strDocumentTitle As String = "", _
    Optional ByRef blnShowPrintDialog As Boolean = False, _
    Optional ByRef blnShowProgressBar As Boolean = False, _
    Optional ByRef strActivePrinter As String = "", _
    Optional ByRef nFirstPagePrinterTray As Integer = wdPrinterDefaultBin, _
    Optional ByRef nOtherPagesPrinterTray As Integer = wdPrinterDefaultBin, _
    Optional ByRef bUseDifferentTrayForOtherPages As Boolean = False, _
    Optional ByRef nCopies As Integer = 1) As Boolean

    g_Log.WriteLine "->PrintWordDocument(" & strFileName & ")"
    
    ' Let caller handle errors.
    
    Const strFunctionName As String = "PrintWordDocument"
    
    Dim bSuccess As Boolean
    
    bSuccess = False
           
    If blnShowPrintDialog Then
        ' NOT USED: Print document using the Word print dialog.
        'pobjWordApp.Dialogs(wdDialogFilePrint).Show
        dlgPrint.DeviceName = strActivePrinter
        dlgPrint.Copies = nCopies
        dlgPrint.FirstPageTray = nFirstPagePrinterTray
        dlgPrint.OtherPagesTray = nOtherPagesPrinterTray
        dlgPrint.UseDifferentTrayForOtherPages = bUseDifferentTrayForOtherPages
        dlgPrint.ShowProgressBar = blnShowProgressBar
        dlgPrint.DocumentTitle = strDocumentTitle

        If Not frmParent Is Nothing Then
            dlgPrint.Show vbModal, frmParent
        Else
'            If App.NonModalAllowed Then
'                ' Not called from Internet Explorer. TaskBar icon.
'                dlgPrint.Show vbModeless
'            Else
'                ' Called from Internet Explorer. No TaskBar icon.
'                dlgPrint.Show vbModal
'            End If
            ' AS 05/01/2005 Must always be modal, otherwise program flow continues immediately
            ' and does not wait for OK/Cancel on dlgPrint.
            dlgPrint.Show vbModal
        End If

        If dlgPrint.OK Then
            strActivePrinter = dlgPrint.DeviceName
            nCopies = dlgPrint.Copies
            nFirstPagePrinterTray = dlgPrint.FirstPageTray
            nOtherPagesPrinterTray = dlgPrint.OtherPagesTray
            bUseDifferentTrayForOtherPages = dlgPrint.UseDifferentTrayForOtherPages
            SavePrintDialogState
            bSuccess = True
        End If
                
    Else
        bSuccess = True
    End If
    
    If bSuccess Then
        
        Dim MessageOp As MessageOpPrintWordDocument
        
        If blnShowProgressBar Then
            Set MessageOp = New MessageOpPrintWordDocument
            frmMessage.MessageOp = MessageOp
            frmMessage.Caption = "Printing"
            If strDocumentTitle = "" Then
                strDocumentTitle = "document"
            End If
            frmMessage.Message = "Please wait - printing " & strDocumentTitle & "."
        End If
        
        gPrintDocumentData.nDeliveryType = DELIVERYTYPE_DOC
        Set gPrintDocumentData.pobjWordApp = pobjWordApp
        gPrintDocumentData.strFileName = strFileName
        gPrintDocumentData.strActivePrinter = strActivePrinter
        gPrintDocumentData.nFirstPagePrinterTray = nFirstPagePrinterTray
        gPrintDocumentData.nOtherPagesPrinterTray = nOtherPagesPrinterTray
        gPrintDocumentData.bUseDifferentTrayForOtherPages = bUseDifferentTrayForOtherPages
        gPrintDocumentData.nCopies = nCopies
        
        If blnShowProgressBar Then
            frmMessage.Show vbModal
            bSuccess = frmMessage.Success
            Set MessageOp = Nothing
        Else
            bSuccess = PrintWordDocumentEx
        End If
            
    End If
        
    PrintWordDocument = bSuccess
    
    g_Log.WriteLine "<-PrintWordDocument(" & strFileName & ") = " & CStr(bSuccess)
    
End Function

Private Sub UpdateProgressBar(ByRef pbProgress As ProgressBar, ByVal Amount As Integer)
    If Not pbProgress Is Nothing Then
        pbProgress = Amount
    End If
End Sub

Public Function PrintWordDocumentEx(Optional ByRef pbProgress As ProgressBar = Nothing) As Boolean
    Dim pobjWordApp As Object
    Dim strFileName As String
    Dim strActivePrinter As String
    Dim nFirstPagePrinterTray As Integer
    Dim nOtherPagesPrinterTray As Integer
    Dim bUseDifferentTrayForOtherPages As Boolean
    Dim nCopies As Integer
    
    g_Log.WriteLine "->PrintWordDocument(" & strFileName & ")"
    
    Set pobjWordApp = gPrintDocumentData.pobjWordApp
    strFileName = gPrintDocumentData.strFileName
    strActivePrinter = gPrintDocumentData.strActivePrinter
    nFirstPagePrinterTray = gPrintDocumentData.nFirstPagePrinterTray
    nOtherPagesPrinterTray = gPrintDocumentData.nOtherPagesPrinterTray
    bUseDifferentTrayForOtherPages = gPrintDocumentData.bUseDifferentTrayForOtherPages
    nCopies = gPrintDocumentData.nCopies
    
    Dim objWordApplication As Object
    Dim objWordDocOut As Object
    Dim blnCreated As Boolean
        
    Const strFunctionName As String = "PrintWordDocumentEx"
        
    Dim bSuccess As Boolean
    bSuccess = False
    
    UpdateProgressBar pbProgress, 10

    If pobjWordApp Is Nothing Then
    
        bSuccess = False
        
        Set objWordApplication = CreateWordApplication(objWordApplication, blnCreated)
        
        
        If Not objWordApplication Is Nothing Then
            UpdateProgressBar pbProgress, 30
            Set pobjWordApp = objWordApplication
            Set objWordDocOut = _
                objWordApplication.Documents.Open( _
                    FileName:=strFileName, _
                    ConfirmConversions:=False, _
                    ReadOnly:=True, _
                    AddToRecentFiles:=False)
            If Not objWordDocOut Is Nothing Then
                UpdateProgressBar pbProgress, 50
                bSuccess = True
            End If
        End If
    End If
    
    If bSuccess Then
        If Len(strActivePrinter) > 0 Then
            ' Set active printer, otherwise use default printer.
            ' Do not use ActivePrinter, as this changes the default printer - see MSDN Q216026.
            pobjWordApp.WordBasic.FilePrintSetup Printer:=strActivePrinter, DoNotSetAsSysDefault:=1
        End If
        
        UpdateProgressBar pbProgress, 60
               
        If pobjWordApp.ActiveDocument.ProtectionType <> wdAllowOnlyComments Then
            ' Set printer tray - defaults to wdPrinterDefaultBin.
            ' Document is not read only, or in non edit mode.
            If nFirstPagePrinterTray <> wdPrinterDefaultBin Then
                pobjWordApp.ActiveDocument.PageSetup.FirstPageTray = nFirstPagePrinterTray
            End If
            If nOtherPagesPrinterTray <> wdPrinterDefaultBin Then
                pobjWordApp.ActiveDocument.PageSetup.OtherPagesTray = nOtherPagesPrinterTray
            End If
        End If
        
        UpdateProgressBar pbProgress, 70
        
        ' Not using Word print dialog, so need to call Word to perform the print.
        ' Note "Backgound:=True" produces garbled print out when using PrintDocument.
        If Not gblnDisablePrintOut Then
            pobjWordApp.ActiveDocument.PrintOut _
                Background:=False, Append:=False, Range:=wdPrintAllDocument, _
                Item:=wdPrintDocumentContent, Copies:=nCopies, PageType:=wdPrintAllPages, _
                PrintToFile:=False, Collate:=True, ManualDuplexPrint:=False
        End If
        
        UpdateProgressBar pbProgress, 80
        
        Debug.Print "Printed Word document to " & strActivePrinter & "; Copies: " & nCopies & "; First page tray: " & nFirstPagePrinterTray & "; Other pages tray: " & nOtherPagesPrinterTray
        g_Log.WriteLine "Printed Word document to " & strActivePrinter & "; Copies: " & nCopies & "; First page tray: " & nFirstPagePrinterTray & "; Other pages tray: " & nOtherPagesPrinterTray
    
        If Not objWordDocOut Is Nothing Then
            Close_Word_Doc objWordDocOut, gstrObjectName
        End If
        If blnCreated Then
            Close_Word_App objWordApplication, gstrObjectName
        End If
        
        UpdateProgressBar pbProgress, 100
        
        gblnDocumentPrinted = True
       
    End If
     
    PrintWordDocumentEx = bSuccess

    g_Log.WriteLine "<-PrintWordDocumentEx(" & strFileName & ") = " & CStr(bSuccess)
End Function

Public Function PrintPDFDocument( _
    ByRef frmParent As Form, _
    ByVal strFileName As String, _
    Optional ByVal strDocumentTitle As String = "", _
    Optional ByVal blnConvertFromRTF As Boolean = False, _
    Optional ByRef blnShowPrintDialog As Boolean = False, _
    Optional ByRef blnShowProgressBar As Boolean = False, _
    Optional ByRef strActivePrinter As String = "", _
    Optional ByRef nFirstPagePrinterTray As Integer = wdPrinterDefaultBin, _
    Optional ByRef nOtherPagesPrinterTray As Integer = wdPrinterDefaultBin, _
    Optional ByRef bUseDifferentTrayForOtherPages As Boolean = False, _
    Optional ByRef nCopies As Integer = 1) As Boolean
    
    ' Let caller handle errors.
    
    g_Log.WriteLine "->PrintPDFDocument(" & strFileName & ")"
    
    Const strFunctionName As String = "PrintPDFDocument"
    
    Dim bSuccess As Boolean
    bSuccess = False
    
    If blnShowPrintDialog Then
        dlgPrint.DeviceName = strActivePrinter
        dlgPrint.Copies = nCopies
        dlgPrint.FirstPageTray = nFirstPagePrinterTray
        dlgPrint.OtherPagesTray = nOtherPagesPrinterTray
        dlgPrint.UseDifferentTrayForOtherPages = bUseDifferentTrayForOtherPages
        dlgPrint.ShowProgressBar = blnShowProgressBar
        dlgPrint.DocumentTitle = strDocumentTitle
        
        If Not frmParent Is Nothing Then
            dlgPrint.Show vbModal, frmParent
        Else
            ' AS 05/01/2005 Must always be modal, otherwise program flow continues immediately
            ' and does not wait for OK/Cancel on dlgPrint.
            dlgPrint.Show vbModal
        End If
        
        If dlgPrint.OK Then
            strActivePrinter = dlgPrint.DeviceName
            nCopies = dlgPrint.Copies
            nFirstPagePrinterTray = dlgPrint.FirstPageTray
            nOtherPagesPrinterTray = dlgPrint.OtherPagesTray
            bUseDifferentTrayForOtherPages = dlgPrint.UseDifferentTrayForOtherPages
            SavePrintDialogState
            bSuccess = True
        End If
    Else
        bSuccess = True
    End If
    
    If bSuccess Then
        
        Dim MessageOp As MessageOpPrintPDFDocument
        
        If blnShowProgressBar Then
            Set MessageOp = New MessageOpPrintPDFDocument
            frmMessage.MessageOp = MessageOp
            frmMessage.Caption = "Printing"
            If strDocumentTitle = "" Then
                strDocumentTitle = "document"
            End If
            frmMessage.Message = "Please wait - printing " & strDocumentTitle & "."
        End If
        
        gPrintDocumentData.nDeliveryType = DELIVERYTYPE_PDF
        gPrintDocumentData.strFileName = strFileName
        gPrintDocumentData.blnConvertFromRTF = blnConvertFromRTF
        gPrintDocumentData.strActivePrinter = strActivePrinter
        gPrintDocumentData.nFirstPagePrinterTray = nFirstPagePrinterTray
        gPrintDocumentData.nOtherPagesPrinterTray = nOtherPagesPrinterTray
        gPrintDocumentData.bUseDifferentTrayForOtherPages = bUseDifferentTrayForOtherPages
        gPrintDocumentData.nCopies = nCopies
                
        If blnShowProgressBar Then
            frmMessage.Show vbModal
            bSuccess = frmMessage.Success
            Set MessageOp = Nothing
        Else
            bSuccess = PrintPDFDocumentEx
        End If
        
    End If
    
    PrintPDFDocument = bSuccess
    
    g_Log.WriteLine "<-PrintPDFDocument(" & strFileName & ") = " & CStr(bSuccess)
End Function

Public Function PrintPDFDocumentEx(Optional ByRef pbProgress As ProgressBar = Nothing) As Boolean
    ' Let caller handle errors.
    
    Dim strFileName As String
    Dim blnConvertFromRTF As Boolean
    Dim strActivePrinter As String
    Dim nFirstPagePrinterTray As Integer
    Dim nOtherPagesPrinterTray As Integer
    Dim bUseDifferentTrayForOtherPages As Boolean
    Dim nCopies As Integer
    
    UpdateProgressBar pbProgress, 10
    
    strFileName = gPrintDocumentData.strFileName
    blnConvertFromRTF = gPrintDocumentData.blnConvertFromRTF
    strActivePrinter = gPrintDocumentData.strActivePrinter
    nFirstPagePrinterTray = gPrintDocumentData.nFirstPagePrinterTray
    nOtherPagesPrinterTray = gPrintDocumentData.nOtherPagesPrinterTray
    bUseDifferentTrayForOtherPages = gPrintDocumentData.bUseDifferentTrayForOtherPages
    nCopies = gPrintDocumentData.nCopies
    
    g_Log.WriteLine "->PrintPDFDocumentEx(" & strFileName & ")"
    
    Const strFunctionName As String = "PrintPDFDocumentEx"
    
    Dim bSuccess As Boolean
    bSuccess = False
    
    UpdateProgressBar pbProgress, 20
    
    Dim strPDFFileName As String
    If blnConvertFromRTF Then
        ' Document is an RTF, so convert to PDF.
        ' This means we do not have to use Word to print the document.
        strPDFFileName = strFileName & ".pdf"
        bSuccess = ConvertRTFToPDF(strFileName, strPDFFileName)
    Else
        strPDFFileName = strFileName
        bSuccess = True
    End If
    
    UpdateProgressBar pbProgress, 50
    
    If bSuccess Then
        Dim nPagesPrinted As Long
        nPagesPrinted = PrintPDFFile(strPDFFileName, strActivePrinter, nFirstPagePrinterTray, nOtherPagesPrinterTray, nCopies)
        bSuccess = nPagesPrinted > 0
        UpdateProgressBar pbProgress, 70
        If blnConvertFromRTF Then
            Call SafeDeleteFile(strPDFFileName)
        End If
        UpdateProgressBar pbProgress, 90
    End If
                
    UpdateProgressBar pbProgress, 100
    
    PrintPDFDocumentEx = bSuccess
    
    g_Log.WriteLine "<-PrintPDFDocumentEx(" & strFileName & ") = " & CStr(bSuccess)
End Function

Private Function PrintPDFFile( _
    ByVal strFileName As String, _
    ByVal strPrinter As String, _
    ByVal nFirstPagePrinterTray As Integer, _
    ByVal nOtherPagesPrinterTray As Integer, _
    Optional ByRef nCopies As Integer = 1) As Long
    
    g_Log.WriteLine "->PrintPDFFile(" & strFileName & ")"
    
    Const strFunctionName As String = "PrintPDFFile"

    Dim nPagesPrinted As Long
    Dim lngLicenceCode As Long
    Dim strPassword As String
    Dim strLicenceName As String
    Dim strLicenceKey As String
    Dim strOptions As String
    Dim strOptionsSeparator As String

    nPagesPrinted = 0
    strLicenceName = "Marlborough-Stirling"
    strLicenceKey = "GAjAd9yU9dk9GA9kyRe"
    lngLicenceCode = 5384401
    strOptionsSeparator = vbCrLf
       
    strOptions = ""
    If Len(strPrinter) > 0 Then
        'AS 01/02/06 CORE234 pdfPrint treats spaces (or commas) in the printer name as delimiters,
        'unless the option is double quoted.
        strOptions = """PRINTER=" & strPrinter & """"
    End If
    
    If Len(strOptions) > 0 Then
        strOptions = strOptions & strOptionsSeparator
    End If
    strOptions = strOptions & "COPIES=" & CStr(nCopies)
    
    If nFirstPagePrinterTray <> wdPrinterDefaultBin Then
        If Len(strOptions) > 0 Then
            strOptions = strOptions & strOptionsSeparator
        End If
        strOptions = strOptions & "TRAY=" & CStr(nFirstPagePrinterTray)
    End If
    
    If nOtherPagesPrinterTray <> wdPrinterDefaultBin Then
        If Len(strOptions) > 0 Then
            strOptions = strOptions & strOptionsSeparator
        End If
        strOptions = strOptions & "TRAY2=" & CStr(nOtherPagesPrinterTray)
    End If
    
    If Len(strOptions) > 0 Then
        strOptions = strOptions & strOptionsSeparator
    End If
    strOptions = strOptions & "QUIET=0"
    If Len(strOptions) > 0 Then
        strOptions = strOptions & strOptionsSeparator
    End If
    strOptions = strOptions & "LISTPRINTER=1"
    If Len(strOptions) > 0 Then
        strOptions = strOptions & strOptionsSeparator
    End If
    strOptions = strOptions & "LISTTRAY=1"
           
    If gblnDisablePrintOut Then
        nPagesPrinted = 1
    Else
        Debug.Print "->pdfPrint(" & strFileName & ", " & strPrinter & ", Copies: " & nCopies & ", " & strOptions & ")"
        g_Log.WriteLine "->pdfPrint(" & strFileName & ", " & strPrinter & ", Copies: " & nCopies & ", " & strOptions & ")"
        nPagesPrinted = pdfPrint(strFileName, strPassword, strLicenceName, strLicenceKey, lngLicenceCode, strOptions)
        Debug.Print "<-pdfPrint(" & strFileName & ", " & strPrinter & ", Copies: " & nCopies & ", " & strOptions & ") = " & nPagesPrinted
        g_Log.WriteLine "<-pdfPrint(" & strFileName & ", " & strPrinter & ", Copies: " & nCopies & ", " & strOptions & ") = " & nPagesPrinted
        If nPagesPrinted > 0 Then
            gblnDocumentPrinted = True
        Else
            Err.Raise vbObjectError, "CommonFunctions.bas::" & strFunctionName, "Error printing " & strFileName & " - pdfPrint returned 0"
        End If
    End If
    
    PrintPDFFile = nPagesPrinted
    
    g_Log.WriteLine "<-PrintPDFFile(" & strFileName & ") = " & CStr(nPagesPrinted)
End Function

Private Function ConvertRTFToPDF(strRTFFileName, strPDFFileName) As String
    On Error GoTo ExitPoint

    g_Log.WriteLine "->ConvertRTFToPDF(" & strRTFFileName & ", " & strPDFFileName & ")"
    
    Const strFunctionName As String = "ConvertRTFToPDF"

    Dim bSuccess As Boolean
    bSuccess = False
    
    Dim objPDF As Object
       
    ' Early binding to wPDF_X01.PDFControl causes compile time errors.
    Set objPDF = CreateObject("wPDF_X01.PDFControl")
    If Not objPDF Is Nothing Then
        Dim strPassword As String
        
        objPDF.PDF_Compression = 1      ' Deflate.
        objPDF.PDF_FontMode = 0
        objPDF.SECURITY_Permission = 0
        objPDF.SECURITY_OwnerPassword = "R3NrstsEs4EdeC7rtff4"
        objPDF.SECURITY_UserPassword = strPassword
        objPDF.INFO_Author = "Marlborough-Stirling"
        objPDF.INFO_Date = Format$(Now(), "dd mmmm yyyy at hh:nn:ss")

        'Dim strSystemDirectory As String
        'Dim uSize As Long
        'uSize = MAX_PATH
        'strSystemDirectory = Space$(uSize)
        'uSize = GetSystemDirectory(strSystemDirectory, uSize)
        'strSystemDirectory = Left(strSystemDirectory, uSize)
        
        Dim strwRTF2PDF02Path As String
        'strwRTF2PDF02Path = strSystemDirectory & "\wRTF2PDF02.dll"
        strwRTF2PDF02Path = App.Path & "\wRTF2PDF02.dll"
        If objPDF.StartEngine(strwRTF2PDF02Path, "Marlborough Stirling plc", "R3NrstsEs4EdeC7rtff4", 16575595) Then
            Dim lResult As Long
            lResult = objPDF.BeginDoc(strPDFFileName, 0)
            If lResult > 0 Then
                objPDF.ExecIntCommand 1024, 1 'Use printer to determine sizes
                objPDF.ExecIntCommand 1000, 0
                objPDF.ExecStrCommand 1002, strRTFFileName
                objPDF.ExecIntCommand 1100, 0
                objPDF.EndDoc
                
                bSuccess = True
            Else
                Err.Raise vbObjectError, "CommonFunctions.bas::" & strFunctionName, "Error converting RTF to PDF - objPDF.BeginDoc returned " & CStr(lResult)
            End If

            objPDF.StopEngine
            Set objPDF = Nothing
        Else
            Err.Raise vbObjectError, "CommonFunctions.bas::" & strFunctionName, "Error printing RTF - unable to load " & strwRTF2PDF02Path
        End If
    Else
        Err.Raise vbObjectError, "CommonFunctions.bas::" & strFunctionName, "Error printing RTF - unable to created wPDF_X01.PDFControl"
    End If

ExitPoint:
    Set objPDF = Nothing
    
    If Err.Number <> 0 Then
        ' Reraise any error.
        Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    End If

    ConvertRTFToPDF = bSuccess

    g_Log.WriteLine "<-ConvertRTFToPDF(" & strRTFFileName & ", " & strPDFFileName & ") = " & CStr(bSuccess)
End Function

Public Function FindFreeTxt(ByVal pstrObjectName As String, _
                            ByRef pobjWordDoc As Object) As Boolean
    ' Find next free text.
    On Error GoTo ExitPoint
    
    Const strFunctionName As String = "FindFreeTxt()"
    
    ' Setup find parameters.
    With pobjWordDoc.ActiveWindow.Selection.Find
        .ClearFormatting
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Text = "...enter free text"
    End With
    
    ' Find operation.
    FindFreeTxt = pobjWordDoc.ActiveWindow.Selection.Find.Execute(Forward:=True)
        
ExitPoint:
    Handle_Error Err, pstrObjectName, strFunctionName, "", False, False
End Function

Public Function ConvertBinToBase64(ByVal pstrObjectName As String, ByRef rarrByte() As Byte, ByRef FileContents As FILE_CONTENTS, ByVal bCompress As Boolean) As Boolean
    ' header ----------------------------------------------------------------------------------
    ' description:
    '   Function to convert binary to base 64.
    '-------------------------------------------------------------------------------------------
    
On Error GoTo ExitPoint
    
    Const strFunctionName = "ConvertBinToBase64"
    
    ConvertBinToBase64 = False
    
    ' Variables.
    Dim xmlDoc As MSXML2.DOMDocument
    Dim xndConversionNode As IXMLDOMNode

    ' Initialise variables.
    Set xmlDoc = New MSXML2.DOMDocument
    If xmlDoc Is Nothing Then
        Err.Raise vbObjectError, strFunctionName, "Failed to create XML Parser, check MSXML is installed."
    End If
    
    ' Compress.
    Dim bCompressed As Boolean
    bCompressed = False
    If bCompress Then
        bCompressed = Compress(rarrByte, FileContents.strCompressionMethod)
        FileContents.bCompressedArray = True
    End If
    
    ' Convert
    Set xndConversionNode = xmlDoc.createElement("CONVERSION_NODE")
    xndConversionNode.dataType = "bin.base64"
    xndConversionNode.nodeTypedValue = rarrByte
    FileContents.strBinBase64 = xndConversionNode.Text
    ' FileContents.strBinBase64 is now compressed.
    FileContents.bCompressed = bCompressed
    ConvertBinToBase64 = True
    
ExitPoint:

    ' Cleanup.
    Set xmlDoc = Nothing
    Set xndConversionNode = Nothing
    
    ' Check for errors.
    Handle_Error Err, pstrObjectName, strFunctionName
End Function

Public Function ConvertBase64ToBin(ByVal pstrObjectName As String, ByRef FileContents As FILE_CONTENTS, ByVal bDecompress As Boolean) As Byte()
' header ----------------------------------------------------------------------------------
' description:
'   Converts base.bin64 to a byte array.
'-------------------------------------------------------------------------------------------
On Error GoTo ExitPoint
    
    Const strFunctionName As String = "ConvertBase64ToBin"
    
    ' Variables.
    Dim xmlDoc As MSXML2.DOMDocument
    Dim xndConversionNode As IXMLDOMNode

    ' Intialise variables.
    Set xmlDoc = New MSXML2.DOMDocument
    If xmlDoc Is Nothing Then
        Err.Raise vbObjectError, strFunctionName, "Failed to create XML Parser, check MSXML is installed."
    End If
    
    ' Convert the bin.base64 stream.
    Set xndConversionNode = xmlDoc.createElement("CONVERSION_NODE")
    xndConversionNode.Text = FileContents.strBinBase64
    xndConversionNode.dataType = "bin.base64"
 
    Dim arrData() As Byte
    arrData = xndConversionNode.nodeTypedValue
    
    ' Decompress.
    If bDecompress And FileContents.bCompressed Then
        Call Decompress(arrData, FileContents.strCompressionMethod)
        ' Do not change FileContents.bCompressed flag as FileContents.strBinBase64 is still
        ' compressed; only the byte array is decompressed, not the original string.
        FileContents.bCompressedArray = False
    End If
        
    ConvertBase64ToBin = arrData
        
ExitPoint:
    ' Clean up.
    Set xmlDoc = Nothing
    Set xndConversionNode = Nothing
    
    ' Check for errors.
    Handle_Error Err, pstrObjectName, strFunctionName
    
End Function

Public Function ConvertBase64ToBase64Compressed(ByVal pstrObjectName As String, ByRef FileContents As FILE_CONTENTS) As Boolean
On Error GoTo ExitPoint
    
    Const strFunctionName As String = "ConvertBase64ToBase64Compressed"
    
    ConvertBase64ToBase64Compressed = False
    If ConvertBinToBase64(pstrObjectName, ConvertBase64ToBin(pstrObjectName, FileContents, True), FileContents, True) Then
        ConvertBase64ToBase64Compressed = True
    End If
    FileContents.bCompressed = True
        
ExitPoint:
    
    ' Check for errors.
    Handle_Error Err, pstrObjectName, strFunctionName
    
End Function

Public Function ConvertBase64ToBase64Decompressed(ByVal pstrObjectName As String, ByRef FileContents As FILE_CONTENTS) As Boolean
On Error GoTo ExitPoint
    
    Const strFunctionName As String = "ConvertBase64ToBase64Decompressed"
    
    ConvertBase64ToBase64Decompressed = False
    If ConvertBinToBase64(pstrObjectName, ConvertBase64ToBin(pstrObjectName, FileContents, True), FileContents, False) Then
        ConvertBase64ToBase64Decompressed = True
    End If
    FileContents.bCompressed = False
        
ExitPoint:
    
    ' Check for errors.
    Handle_Error Err, pstrObjectName, strFunctionName
    
End Function

Private Function Compress(ByRef arrData() As Byte, ByVal strCompressionMethod As String) As Boolean
    
    Compress = False
    
    If Len(strCompressionMethod) > 0 Then
        Dim objCompression As dmsCompression1
        Set objCompression = New DMSCOMPRESSIONLib.dmsCompression1
        
        If Not objCompression Is Nothing Then
            objCompression.CompressionMethod = strCompressionMethod
            arrData = objCompression.SafeArrayCompressToSafeArray(arrData)
            Set objCompression = Nothing
        End If
    End If
    
    Compress = True
    
End Function

Private Function Decompress(ByRef arrData() As Byte, ByVal strCompressionMethod As String) As Boolean
    Decompress = False
    
    If Len(strCompressionMethod) > 0 Then
        Dim objCompression As dmsCompression1
        Set objCompression = New DMSCOMPRESSIONLib.dmsCompression1
        
        If Not objCompression Is Nothing Then
            objCompression.CompressionMethod = strCompressionMethod
            arrData = objCompression.SafeArrayDecompressToSafeArray(arrData)
            Set objCompression = Nothing
        End If
    End If
    
    Decompress = True
End Function

Public Function Create_Word_App(ByVal pstrObjectName As String) As Object
    ' Create a word application.
    On Error GoTo ExitPoint

    Const strFunctionName As String = "Create_Word_App"
        
    ' Create application.
    Set Create_Word_App = CreateObject("Word.Application")
    
    Create_Word_App.Options.SaveNormalPrompt = False
    Create_Word_App.NormalTemplate.Saved = True
    
    glWordVersion = getWordVersion(Create_Word_App)
    
ExitPoint:
    Handle_Error Err, pstrObjectName, strFunctionName
    
End Function


Public Sub Close_Word_App(ByRef pobjWordApp As Object, _
                           ByVal pstrObjectName As String)
    ' Close a word application.
    On Error GoTo ExitPoint

    Const strFunctionName As String = "Close_Word_App"
        
    ' Close the application if required.
    If (Not pobjWordApp Is Nothing) Then
        'Stop complaining about the template
        pobjWordApp.NormalTemplate.Saved = True
        ' Quit and clear object.
        pobjWordApp.Quit SaveChanges:=wdDoNotSaveChanges
        Set pobjWordApp = Nothing
    End If
    
ExitPoint:
    Handle_Error Err, pstrObjectName, strFunctionName
    
End Sub


Public Sub Close_Word_Doc(ByRef pobjWordDoc As Object, _
                           ByVal pstrObjectName As String)
    ' Close a word application.
    On Error GoTo ExitPoint

    Const strFunctionName As String = "Close_Word_Doc"
        
    ' Close the application if required.
    If (Not pobjWordDoc Is Nothing) Then
        ' Close all documents supressing the save changes dialogue.
        pobjWordDoc.Close SaveChanges:=wdDoNotSaveChanges
        
        Set pobjWordDoc = Nothing
    End If
    
ExitPoint:
    Handle_Error Err, pstrObjectName, strFunctionName
    
End Sub


Public Sub Handle_Error(ByRef pobjErr As ErrObject, _
                         ByVal pstrObjectName As String, _
                         ByVal pstrFunctionName As String, _
                         Optional ByVal pstrMessage As String = "", _
                         Optional ByVal blnRaiseError As Boolean = True, _
                         Optional ByVal blnExitOnError As Boolean = True, _
                         Optional ByVal blnDisplayError As Boolean = False)
    ' Report error and close application.
    Dim strError As String
    
    ' Raise error if required.
    If pobjErr.Number <> 0 Or Len(pstrMessage) > 0 Then
        'Set up the error details
        If (pobjErr.Number <> 0) Then
             strError = "Error in " & pstrObjectName & "::" & pstrFunctionName & vbCrLf & _
                        "Error Code: " & pobjErr.Number & vbCrLf & _
                        "Error Source: " & pobjErr.Source & vbCrLf & _
                        "Error Description: " & pobjErr.Description
            If Len(pstrMessage) > 0 Then
                strError = strError & vbCrLf & "Additional Info: " & pstrMessage
            End If
            gLastErrObject.Number = pobjErr.Number
            gLastErrObject.Source = pobjErr.Source
            gLastErrObject.Description = pobjErr.Description
        Else
            strError = "Error in " & pstrObjectName & "::" & pstrFunctionName
            If Len(pstrMessage) > 0 Then
                 strError = strError & vbCrLf & "Error Description: " & pstrMessage
            End If
            gLastErrObject.Number = vbObjectError
            gLastErrObject.Source = pstrObjectName & "::" & pstrFunctionName
            gLastErrObject.Description = pstrMessage
        End If

        'Log in the event log
        App.LogEvent strError, vbLogEventTypeError
        
#If Not MIN_BUILD Then
        If gblnWindowShown Then
            'Display to user
            frmaxword.MousePointer = vbDefault
            
            MsgBox strError, vbCritical, "Error"
             
            ' Unload form if required.
            If blnExitOnError Then
                frmaxword.Enabled = True
                frmaxword.cmdExit.Enabled = True
            End If
        ElseIf blnDisplayError Then
            MsgBox strError, vbCritical, "Error"
        End If
#End If

        ' Output debug message.
        Debug.Print strError
        
        ' Set the global error flag.
        gblnError = True
        
#If Not MIN_BUILD Then
        If blnExitOnError Then
            ' Quit the app if we are not already trying to.
            If gblnInError = False Then
                gblnInError = True
                If gblnWindowShown Then
                    Unload frmaxword
                End If
            End If
        End If
#End If

        If blnRaiseError Then
            ' Re-raise error.
            If pobjErr.Number <> 0 Then
                Err.Raise pobjErr.Number, pobjErr.Source, pobjErr.Description
            Else
                Err.Raise vbObjectError, pstrFunctionName, pstrMessage
            End If
        End If
    End If
End Sub

Public Sub LoadPrintDialogState()
    Dim strSection As String
    strSection = "Axword\Printer"
    
    If gRequestData.strDocumentID <> "" Then
        ' Load default settings for this document.
        LoadPrintDialogStateSection strSection + "\" + gRequestData.strDocumentID
        If gRequestData.strPrinter = "" Then
            ' No entries for this document, so load default settings for all documents.
            LoadPrintDialogStateSection strSection
        End If
    Else
        LoadPrintDialogStateSection strSection
    End If
    
End Sub

Private Sub LoadPrintDialogStateSection(ByVal strSection As String)
    gRequestData.strPrinter = GetSetting("Marlborough Stirling", strSection, "DeviceName", gRequestData.strPrinter)
    gRequestData.nCopies = GetSetting("Marlborough Stirling", strSection, "Copies", gRequestData.nCopies)
    gRequestData.nFirstPagePrinterTray = GetSetting("Marlborough Stirling", strSection, "FirstPageTray", gRequestData.nFirstPagePrinterTray)
    gRequestData.nOtherPagesPrinterTray = GetSetting("Marlborough Stirling", strSection, "OtherPagesTray", gRequestData.nOtherPagesPrinterTray)
    gRequestData.bUseDifferentTrayForOtherPages = GetSetting("Marlborough Stirling", strSection, "UseDifferentTrayForOtherPages", gRequestData.bUseDifferentTrayForOtherPages)
End Sub

Public Sub SavePrintDialogState()
    Dim strSection As String
    strSection = "Axword\Printer"
    
    If gRequestData.strDocumentID <> "" Then
        ' Save default settings for this document.
        strSection = strSection + "\" + gRequestData.strDocumentID
    End If
    
    SavePrintDialogStateSection strSection
End Sub

Private Sub SavePrintDialogStateSection(ByVal strSection As String)
    SaveSetting "Marlborough Stirling", strSection, "DeviceName", gRequestData.strPrinter
    SaveSetting "Marlborough Stirling", strSection, "Copies", gRequestData.nCopies
    SaveSetting "Marlborough Stirling", strSection, "FirstPageTray", gRequestData.nFirstPagePrinterTray
    SaveSetting "Marlborough Stirling", strSection, "OtherPagesTray", gRequestData.nOtherPagesPrinterTray
    SaveSetting "Marlborough Stirling", strSection, "UseDifferentTrayForOtherPages", gRequestData.bUseDifferentTrayForOtherPages
End Sub


Public Function GetPrinters(ByRef PrinterData() As PRINTER_DATA, Optional ByRef pbProgress As ProgressBar = Nothing) As Boolean
On Error GoTo errorPoint
    Dim bSuccess As Boolean
    bSuccess = False

    g_Log.WriteLine "->GetPrinters"

    Dim nPrinter As Integer
    nPrinter = 0
    
    ReDim PrinterData(0 To Printers.Count - 1)

    Dim strDefaultPrinterDeviceName As String
    strDefaultPrinterDeviceName = Printer.DeviceName

    Dim currentProgress As Integer
    Dim progressIncrement As Integer

    currentProgress = 0
    progressIncrement = 100 / Printers.Count

    g_Log.WriteLine "Printers.Count = " & CStr(Printers.Count)

    Dim thisPrinter As Printer
    For Each thisPrinter In Printers
        Dim nBins As Long
        Dim strBinNames As String
        Dim DM As DEVMODE
        Dim strDeviceName As String
        Dim strPortName As String

        On Error Resume Next

        strDeviceName = thisPrinter.DeviceName
        strPortName = thisPrinter.Port

        PrinterData(nPrinter).strDeviceName = strDeviceName

        g_Log.WriteLine "PrinterData(" & CStr(nPrinter) & ").strDeviceName = " & PrinterData(nPrinter).strDeviceName

        nBins = DeviceCapabilities(strDeviceName, strPortName, DC_BINS, ByVal vbNullString, DM)

        g_Log.WriteLine "PrinterData(" & CStr(nPrinter) & ") Bins = " & CStr(nBins)

        If nBins > 0 Then
            ReDim PrinterData(nPrinter).strBinNames(0 To nBins)
            ReDim PrinterData(nPrinter).nBinNumbers(0 To nBins)
            'PrinterData(nPrinter).nDefaultBin = DM.dmDefaultSource
            PrinterData(nPrinter).nDefaultBin = wdPrinterDefaultBin
            g_Log.WriteLine "PrinterData(" & CStr(nPrinter) & ").nDefaultBin = " & CStr(wdPrinterDefaultBin)

            strBinNames = String(24 * nBins, 0)
            nBins = DeviceCapabilities(strDeviceName, strPortName, DC_BINS, PrinterData(nPrinter).nBinNumbers(1), DM)
            nBins = DeviceCapabilities(strDeviceName, strPortName, DC_BINNAMES, ByVal strBinNames, DM)
            PrinterData(nPrinter).nBins = nBins + 1
            g_Log.WriteLine "PrinterData(" & CStr(nPrinter) & ").nBins = " & CStr(PrinterData(nPrinter).nBins)

            Dim nBin As Long
            nBin = 0
            PrinterData(nPrinter).strBinNames(nBin) = "Default"
            g_Log.WriteLine "PrinterData(" & CStr(nPrinter) & ").strBinNames(" & CStr(nBin) & ") = " & PrinterData(nPrinter).strBinNames(nBin)

            PrinterData(nPrinter).nBinNumbers(nBin) = PrinterData(nPrinter).nDefaultBin
            g_Log.WriteLine "PrinterData(" & CStr(nPrinter) & ").nBinNumbers(" & CStr(nBin) & ") = " & PrinterData(nPrinter).nDefaultBin

            For nBin = 1 To nBins
                Dim strBinName As String
                strBinName = Mid(strBinNames, 24 * (nBin - 1) + 1, 24)
                strBinName = Left(strBinName, InStr(1, strBinName, Chr(0)) - 1)
                PrinterData(nPrinter).strBinNames(nBin) = strBinName
                g_Log.WriteLine "PrinterData(" & CStr(nPrinter) & ").strBinNames(" & CStr(nBin) & ") = " & PrinterData(nPrinter).strBinNames(nBin)
                g_Log.WriteLine "PrinterData(" & CStr(nPrinter) & ").nBinNumbers(" & CStr(nBin) & ") = " & PrinterData(nPrinter).nBinNumbers(nBin)
            Next nBin
        End If

        PrinterData(nPrinter).bDefault = False
        If PrinterData(nPrinter).strDeviceName = strDefaultPrinterDeviceName Then
            PrinterData(nPrinter).bDefault = True
        End If
        g_Log.WriteLine "PrinterData(" & CStr(nPrinter) & ").bDefault = " & CStr(PrinterData(nPrinter).bDefault)

        nPrinter = nPrinter + 1

        currentProgress = IIf(currentProgress + progressIncrement < 100, currentProgress + progressIncrement, 100)
        UpdateProgressBar pbProgress, currentProgress

    Next thisPrinter

    On Error GoTo errorPoint
    
    bSuccess = nPrinter > 0
    
ExitPoint:
    GetPrinters = bSuccess

    WritePrintersToLog PrinterData
    
    g_Log.WriteLine "<-GetPrinters = " & CStr(bSuccess)
    
    Exit Function
    
errorPoint:
    g_Log.WriteLine "Error in GetPrinters: " & Err.Description
    
    GoTo ExitPoint
    
End Function

Public Function WritePrintersToLog(ByRef PrinterData() As PRINTER_DATA)
On Error GoTo errorPoint

    WritePrintersToLog = False

    If Not g_Log Is Nothing Then
        If g_Log.EnableLog Then
            g_Log.WriteLine "->WritePrintersToLog"
            
            Dim nPrinter As Integer
            For nPrinter = 0 To UBound(PrinterData)
                g_Log.WriteLine "PrinterData(" & CStr(nPrinter) & ").strDeviceName = " & PrinterData(nPrinter).strDeviceName
                g_Log.WriteLine "PrinterData(" & CStr(nPrinter) & ").bDefault = " & CStr(PrinterData(nPrinter).bDefault)
                g_Log.WriteLine "PrinterData(" & CStr(nPrinter) & ").nBins = " & CStr(PrinterData(nPrinter).nBins)
                g_Log.WriteLine "PrinterData(" & CStr(nPrinter) & ").nDefaultBin = " & CStr(PrinterData(nPrinter).nDefaultBin)
                
                Dim nBin As Integer
                For nBin = 0 To UBound(PrinterData(nPrinter).nBinNumbers)
                    g_Log.WriteLine "PrinterData(" & CStr(nPrinter) & ").nBinNumbers(" & nBin & ") = " & CStr(PrinterData(nPrinter).nBinNumbers(nBin))
                Next
                For nBin = 0 To UBound(PrinterData(nPrinter).strBinNames)
                    g_Log.WriteLine "PrinterData(" & CStr(nPrinter) & ").strBinNames(" & nBin & ") = " & PrinterData(nPrinter).strBinNames(nBin)
                Next
            Next
            
            g_Log.WriteLine "<-WritePrintersToLog = " & CStr(True)
        End If
    End If

    WritePrintersToLog = True
    
ExitPoint:
    Exit Function
    
errorPoint:
    g_Log.WriteLine "Error in WritePrintersToLog: " & Err.Description
    
    GoTo ExitPoint
    
End Function

'Public Function GetPrintersNew(ByRef PrinterData() As PRINTER_DATA, Optional ByRef pbProgress As ProgressBar = Nothing) As Boolean
'On Error GoTo ExitPoint
'    Dim bSuccess As Boolean
'    bSuccess = False
'
'    Dim nPrinter As Integer
'    nPrinter = 0
'
'    Dim strDefaultPrinterDeviceName As String
'    'FIXAS
'    strDefaultPrinterDeviceName = ""
'
'    Dim currentProgress As Integer
'    Dim progressIncrement As Integer
'
'    currentProgress = 0
'    progressIncrement = 100 / Printers.Count
'
'    Dim lpPrnInfo As PRINTER_INFO_2
'
'    Dim lngRet As Long
'    Dim lpBuffer() As Byte
'    Dim lpcb As Long
'    Dim lpcbNeeded As Long
'    Dim lpcReturned As Long
'    Dim i As Long
'    Dim ppPrnInfo As Long
'    lpcb = 4096
'    ReDim lpBuffer(lpcb)
'    lngRet = EnumPrinters(PRINTER_ENUM_CONNECTIONS Or PRINTER_ENUM_LOCAL, vbNullString, 2, lpBuffer(0), lpcb, lpcbNeeded, lpcReturned)
'    If lngRet = 0 Then
'        Erase lpBuffer
'        lpcbNeeded = lpcbNeeded + 1024
'        ReDim lpBuffer(lpcbNeeded)
'        lpcb = lpcbNeeded
'        lngRet = EnumPrinters(PRINTER_ENUM_CONNECTIONS Or PRINTER_ENUM_LOCAL, vbNullString, 2, lpBuffer(0), lpcb, lpcbNeeded, lpcReturned)
'    End If
'
'    If Not lngRet = 0 Then
'        ppPrnInfo = VarPtr(lpBuffer(0))
'
'        ReDim PrinterData(0 To lpcReturned - 1)
'
'        For i = 0 To lpcReturned - 1
'            Dim nBins As Long
'            Dim strBinNames As String
'            Dim DM As DEVMODE
'            Dim strDeviceName As String
'            Dim strPortName As String
'
'            On Error Resume Next
'
'            Call CopyMemory(lpPrnInfo, ByVal ppPrnInfo, Len(lpPrnInfo))
'
'            strDeviceName = Space$(lstrlenA(lpPrnInfo.pPrinterName) + 1)
'            Call CopyMemory(ByVal strDeviceName, ByVal lpPrnInfo.pPrinterName, Len(strDeviceName) - 1)
'            strDeviceName = Left(strDeviceName, Len(strDeviceName) - 1)
'            PrinterData(nPrinter).strDeviceName = strDeviceName
'
'            strPortName = Space$(lstrlenA(lpPrnInfo.pPortName) + 1)
'            Call CopyMemory(ByVal strPortName, ByVal lpPrnInfo.pPortName, Len(strPortName) - 1)
'            strPortName = Left(strPortName, Len(strPortName) - 1)
'
'            nBins = DeviceCapabilities(strDeviceName, strPortName, DC_BINS, ByVal vbNullString, DM)
'            If nBins > 0 Then
'                ReDim PrinterData(nPrinter).strBinNames(0 To nBins)
'                ReDim PrinterData(nPrinter).nBinNumbers(0 To nBins)
'                'PrinterData(nPrinter).nDefaultBin = DM.dmDefaultSource
'                PrinterData(nPrinter).nDefaultBin = wdPrinterDefaultBin
'                strBinNames = String(24 * nBins, 0)
'                nBins = DeviceCapabilities(strDeviceName, strPortName, DC_BINS, PrinterData(nPrinter).nBinNumbers(1), DM)
'                nBins = DeviceCapabilities(strDeviceName, strPortName, DC_BINNAMES, ByVal strBinNames, DM)
'                PrinterData(nPrinter).nBins = nBins + 1
'
'                Dim nBin As Long
'                nBin = 0
'                PrinterData(nPrinter).strBinNames(nBin) = "Default"
'                PrinterData(nPrinter).nBinNumbers(nBin) = PrinterData(nPrinter).nDefaultBin
'                For nBin = 1 To nBins
'                    Dim strBinName As String
'                    strBinName = Mid(strBinNames, 24 * (nBin - 1) + 1, 24)
'                    strBinName = Left(strBinName, InStr(1, strBinName, Chr(0)) - 1)
'                    PrinterData(nPrinter).strBinNames(nBin) = strBinName
'                Next nBin
'            End If
'
'            PrinterData(nPrinter).bDefault = False
'            If PrinterData(nPrinter).strDeviceName = strDefaultPrinterDeviceName Then
'                PrinterData(nPrinter).bDefault = True
'            End If
'
'            ppPrnInfo = ppPrnInfo + Len(lpPrnInfo)
'
'            nPrinter = nPrinter + 1
'
'            currentProgress = IIf(currentProgress + progressIncrement < 100, currentProgress + progressIncrement, 100)
'            UpdateProgressBar pbProgress, currentProgress
'
'        Next
'    End If
'
'    bSuccess = nPrinter > 0
'
'ExitPoint:
'    GetPrintersNew = bSuccess
'End Function

Public Function SafeDeleteFile(ByVal strFileName As String) As Boolean
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String
    Dim errHelpFile As String
    Dim errHelpContext As String
    
    ' Save Err object as it will be cleared by this function.
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    errHelpFile = Err.HelpFile
    errHelpContext = Err.HelpContext
    
    On Error GoTo ExitPoint
    
    SafeDeleteFile = False
    Kill strFileName
    SafeDeleteFile = True

ExitPoint:

    If errNumber <> 0 Then
        ' Restore Err object.
        Err.Number = errNumber
        Err.Source = errSource
        Err.Description = errDescription
        Err.HelpFile = errHelpFile
        Err.HelpContext = errHelpContext
        
    End If
    
End Function

