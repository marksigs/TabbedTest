VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmaxword 
   BackColor       =   &H80000004&
   Caption         =   "Active X Word Viewer"
   ClientHeight    =   7905
   ClientLeft      =   4065
   ClientTop       =   2655
   ClientWidth     =   9120
   FillColor       =   &H80000004&
   ForeColor       =   &H8000000D&
   Icon            =   "frmaxword.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   527
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   608
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   375
      Left            =   3360
      TabIndex        =   6
      Top             =   7440
      Width           =   975
   End
   Begin VB.Frame FrameOptions 
      Caption         =   "Options"
      Height          =   615
      Left            =   120
      TabIndex        =   8
      Top             =   6720
      Width           =   8895
      Begin VB.CheckBox CheckTrackedChanges 
         Caption         =   "Show &tracked changes"
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton cmdFindFreeTxt 
         Caption         =   "&Find Free Text"
         Height          =   375
         Left            =   7560
         TabIndex        =   2
         Top             =   180
         Width           =   1215
      End
   End
   Begin VB.TextBox txtStatus 
      BackColor       =   &H80000004&
      CausesValidation=   0   'False
      Enabled         =   0   'False
      ForeColor       =   &H80000007&
      Height          =   285
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Text            =   "Please wait while initialising..."
      Top             =   7440
      Width           =   4575
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   7440
      Width           =   975
   End
   Begin SHDocVwCtl.WebBrowser web1 
      CausesValidation=   0   'False
      Height          =   6705
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9135
      ExtentX         =   16113
      ExtentY         =   11827
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   0
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Top             =   7440
      Width           =   975
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   2280
      TabIndex        =   5
      Top             =   7440
      Width           =   975
   End
End
Attribute VB_Name = "frmaxword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Workfile:      frmaxword.frm
'Copyright:     Copyright © 2002 Marlborough Stirling

'Description:   Users GUI interface to view files from within the DMS2 system
'Dependencies:  axwordclass.cls (closely linked to this class)
'Issues:
'
'------------------------------------------------------------------------------------------
'History:
'
'Prog   Date     Description
'DR     06/03/02 Created
'DJB    08/05/02 Many issues resolved including Word activation.
'DJB    09/08/02 Converted to process Word files instead of Word documents.
'DJB    02/10/02 Changed OLE embedding to Active X documents, switched to late binding, and
'                added a tracked changes switch.
'DJB    14/10/02 Fix to stop focus exception when the user hasn't clicked on the web object in edit mode.
'LD     15/10/02 Word 97 compatible
'DJB    16/10/02 Removed page size restoration.
'DJB    14/02/03 Added 'Find Free Text' Button and reenabled the form if an error occurs to stop reported
'                hang in some error scenarios.
'DJB    21/05/03 Added logic to retry for upto 60 seconds to gain a lock on a file, this should resolve
'                the permission denied error.
'DJB    17/07/03 Change to make window always on top.
'TW     16/06/04 Modified to handle pdf
'------------------------------------------------------------------------------------------
Option Explicit

#If Not MIN_BUILD Then
Public mobjCurrentDocument As Object
Public mstrStartupStatus As String

'File name - no extension just the name of the file.
'Its public so that it can be changed externally by the calling class.
Public gstrFile As String

'Used in the error reporting to tell us where the error has occured.
Private Const gstrObjectName = "AxwordClass.cls"


Private Sub cmdFindFreeTxt_Click()
    ' Find next free text.
    On Error GoTo ExitPoint
    
    Const strFunctionName As String = "cmdFindFreeTxt_Click()"
            
    ' Find operation.
    If (FindFreeTxt(gstrObjectName, mobjCurrentDocument) = True) Then
        ' Change the text colour to black.
        ' AS 31/03/04 ColorIndex works with Word 97.
        mobjCurrentDocument.ActiveWindow.Selection.Font.ColorIndex = wdBlack
        
        ' Set the focus back to word.
        web1.SetFocus
    Else
        ' No free format text blocks found.
        MsgBox "No free format text blocks found in document.", vbInformation, "Information"
    End If
        
ExitPoint:
    Handle_Error Err, gstrObjectName, strFunctionName, "", False, False
End Sub

Private Sub CheckTrackedChanges_Click()
    On Error GoTo ExitPoint
    
    MousePointer = vbHourglass
    
    Const strFunctionName As String = "CheckTrackedChanges_Click()"
    
    If gblnViewAsWord Then
        If gblnShowTrackedChanges Then
            gblnShowRevisions = IIf(CheckTrackedChanges.Value = 1, True, False)
        Else
            gblnShowRevisions = False
        End If
        mobjCurrentDocument.ShowRevisions = gblnShowRevisions
    Else
        ' Turn on and off visible revisions.
        If gblnShowTrackedChanges And CheckTrackedChanges.Value = 1 Then
            gblnShowRevisions = True
            If gblnEdit Then
                mobjCurrentDocument.ShowRevisions = True
            End If
        Else
            gblnShowRevisions = False
            If gblnEdit Then
                mobjCurrentDocument.ShowRevisions = False
            End If
        End If
        
        ' Recreate HTML if required.
        If Not gblnEdit Then
            CheckTrackedChanges.Enabled = False
            txtStatus.Text = "Generating new preview."
            DoEvents
            
            saveBinBase64AsWordDoc gRequestData.FileContents, gRequestData.nDeliveryType, gstrFile, gblnViewAsWord
            
            ' Refresh browser.
            web1.Refresh
            DoEvents
            txtStatus.Text = "Ready."
            CheckTrackedChanges.Enabled = True
        End If
    End If
    
ExitPoint:
    
    MousePointer = vbDefault
    
    Handle_Error Err, gstrObjectName, strFunctionName, "", False, False
End Sub

Public Sub cmdEdit_Click()
    On Error GoTo ExitPoint
    
    'Used by the error handling routines
    Const strFunctionName As String = "cmdEdit_Click"
      
    ' Update status.
    frmaxword.txtStatus.Text = "Preparing to edit document."
    DoEvents
 
    If (Len(gstrFile) > 0) Then
        'Its OK to display the word doc version of the document for editing
        MousePointer = vbHourglass
        
        'DR We may have an issue with creating the embed file when the next line
        'of code is called. So to be on the safe side we'll control the error handling
        'manually - so that we can report a meaningful error message to the user.
        On Error Resume Next

        If Not gblnReadOnly Then
            gblnEdit = True
        End If
        
        If Not mobjCurrentDocument Is Nothing Then
            mobjCurrentDocument.Saved = True
            web1.Navigate "about:blank"
        End If

        web1.Navigate gstrFile & gRequestData.strFileExtension

        DoEvents
             
        If Err.Number <> 0 Then
            'Right we had a problem, throw a meaningfull error.
            Handle_Error Err, gstrObjectName, strFunctionName, "Could not load the Word document, probably because the file does not exist. File name: " & gstrFile & gRequestData.strFileExtension
        End If
        
        'Clear the error and then turn the error handler back on.
        On Error GoTo ExitPoint
    End If

ExitPoint:
    Handle_Error Err, gstrObjectName, strFunctionName, "", False, False
    
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    On Error GoTo ExitPoint
    
    'Used by the error handling routines
    Const strFunctionName As String = "cmdSave_Click"
    
    ' Variables.
    Dim arrFile() As Byte
    Dim objFile As Object
    Dim intTries As Integer
    
    ' Disable form for safety.
    cmdEdit.Enabled = False
    cmdExit.Enabled = False
    cmdSave.Enabled = False
    cmdFindFreeTxt.Enabled = False
    CheckTrackedChanges.Enabled = False
    
    ' Focus on Word.
    web1.SetFocus
    
    DoEvents
    
    ' Check spelling if required.
    If gblnSpellCheckOnSave And (mobjCurrentDocument.SpellingErrors.Count > 0 Or (getWordVersion(mobjCurrentDocument.Application) >= 8)) Then
        On Error Resume Next ' Required due to exception error.
        mobjCurrentDocument.CheckSpelling
        Debug.Print "Word returned " & Err.Number & " after checking spelling."
        On Error GoTo ExitPoint
        
        DoEvents
    End If
            
    ' Save doc
    MousePointer = vbHourglass
    
    ' Update status.
    frmaxword.txtStatus.Text = "Saving changes"
    DoEvents
           
    SaveCurrentWordDocument
        
    gblnDocumentEdited = True

    If gblnViewAsWord Then
        mobjCurrentDocument.Protect wdAllowOnlyComments
    Else
        'Close word doc and application.
        Set mobjCurrentDocument = Nothing
        web1.Navigate "about:blank"
    End If

    DoEvents
    
    'Convert the word document into Base64
    Dim strFileName As String
    Dim objFSO As FileSystemObject
    Set objFSO = New FileSystemObject
    If objFSO Is Nothing Then
        Err.Raise vbObjectError, strFunctionName, "Failed to create FileSystemObject, check Windows Scripting Host installation."
    End If
    If gblnViewAsWord Then
        ' If viewing as Word document then work on a copy of the original file.
        strFileName = gstrFile & "_wip" & gRequestData.strFileExtension
        ' FileSystemObject.CopyFile can copy an open file, whereas VB FileCopy can not.
        objFSO.CopyFile gstrFile & gRequestData.strFileExtension, strFileName, True
    Else
        strFileName = gstrFile & gRequestData.strFileExtension
    End If
    Set objFSO = Nothing
        
    ' Load the file into memory.
    On Error Resume Next
        For intTries = 1 To gintFileRetries
            Open strFileName For Binary Access Read Lock Write As #1
            If Err.Number <> 0 Then
                ' File still locked, sleep and try again.
                Sleep 500
            Else
                Exit For
            End If
        Next
        
        ' Check for a persistant error.
        If (Err.Number <> 0) Then
            GoTo ExitPoint
        End If
    On Error GoTo ExitPoint
        
    Dim lFileLength As Long
    lFileLength = FileLen(strFileName)
    If gRequestData.nDeliveryType = DELIVERYTYPE_RTF Then
        ' RTF must be read in as text, not binary.
        Dim strFile As String
        strFile = Space(lFileLength)
        Get #1, , strFile
        Debug.Print " " & Asc(Mid(strFile, 1, 1)) & " " & Asc(Mid(strFile, 2, 1)) & " " & Asc(Mid(strFile, 3, 1)) & " " & Asc(Mid(strFile, 4, 1)) & " " & Asc(Mid(strFile, 5, 1)) & " " & Asc(Mid(strFile, 6, 1))
        arrFile = strFile
    Else
        ReDim arrFile(lFileLength - 1)
        Get #1, , arrFile
    End If
    Close #1
    
    'Update the file contents, converting to bin base 64; do not compress.
    If ConvertBinToBase64(gstrObjectName, arrFile, gRequestData.FileContents, bCompress:=False) Then
        If Not gblnViewAsWord Then
            'Need to save the new XML as a HTML file. This will decompress the file contents if necessary.
            saveBinBase64AsWordDoc gRequestData.FileContents, gRequestData.nDeliveryType, gstrFile, gblnViewAsWord
            'Swap back to viewing the HTML version of the doc in the Web object
            web1.Navigate gstrFile & ".htm"
        End If
    End If
    
    'Update Global Variables
    gFileSaved = True
    gblnEdit = False
    MousePointer = vbDefault

    DoEvents
    
    'Update GUI Buttons
    cmdEdit.Enabled = True
    cmdExit.Enabled = True
    
    If gblnShowTrackedChanges And gblnTrackedChanges Then
        CheckTrackedChanges.Enabled = True
    End If
    
    ' Update status.
    frmaxword.txtStatus.Text = "Ready."
    
    DoEvents
 
ExitPoint:
    Set objFSO = Nothing

    Handle_Error Err, gstrObjectName, strFunctionName, "", False, False
    
End Sub

Private Function SaveCurrentWordDocument() As Boolean
    If Not mobjCurrentDocument.Saved Then
    
        ' Stop complaining about the template.
        mobjCurrentDocument.Application.NormalTemplate.Saved = True
        
        ' AS 24/03/04 mobjCurrentDocument.Save with Word 2002 gives the error 4605:
        ' "The Save method or property is not available because this document is in
        ' another application".
        ' This was raised with Microsoft as SRQ040317600632 who acknowledged it is a
        ' bug in Word 2002 and suggested the following workaround.
        web1.ExecWB OLECMDID_SAVE, OLECMDEXECOPT_DONTPROMPTUSER
        
        If gRequestData.nDeliveryType = DELIVERYTYPE_RTF Then
            ' AS 03/02/05 If web1.ExecWB OLECMDID_SAVE is used to save an RTF document, Word
            ' automatically converts the document to Word format, whereas we need to keep it in
            ' RTF format. We can't use web1.ExecWB OLECMDID_SAVEAS as this prompts the user, and
            ' there seems to be no way of turning off the prompt.
            ' Therefore, use Word Document.SaveAs, although this probably will not work with
            ' Word 2002 (see below); it seems to be OK with Word 2003.
            ' Note, we still need to save the document using OLECMDID_SAVE above, otherwise Word
            ' will think the document has not been saved (despite using SaveAs), and will prompt
            ' the user to save the document on quitting Axword.
            ' Setting mobjCurrentDocument.Saved = True has no effect.
            Dim strFullPath As String
            strFullPath = mobjCurrentDocument.FullName
            mobjCurrentDocument.SaveAs FileName:=strFullPath, FileFormat:=wdFormatRTF, AddToRecentFiles:=False
        End If
    End If
        
    SaveCurrentWordDocument = True
End Function

Private Sub cmdPrint_Click()
    On Error GoTo ExitPoint
    
    'Used by the error handling routines
    Const strFunctionName As String = "cmdPrint_Click"
    
    Dim strStatus As String
    strStatus = txtStatus.Text
    txtStatus.Text = "Printing document"
    
    If Not gblnShowPrintDialog Then
        MousePointer = vbHourglass
    End If
    
    If gblnViewAsPDF Then
        PrintPDFDocument frmaxword, gstrFile & ".pdf", gRequestData.strDocumentTitle, False, gblnShowPrintDialog, gblnShowProgressBar, gRequestData.strPrinter, gRequestData.nFirstPagePrinterTray, gRequestData.nOtherPagesPrinterTray, gRequestData.bUseDifferentTrayForOtherPages, gRequestData.nCopies
    ElseIf gblnViewAsWord Then
        PrintWordDocument frmaxword, mobjCurrentDocument.Application, "", gRequestData.strDocumentTitle, gblnShowPrintDialog, gblnShowProgressBar, gRequestData.strPrinter, gRequestData.nFirstPagePrinterTray, gRequestData.nOtherPagesPrinterTray, gRequestData.bUseDifferentTrayForOtherPages, gRequestData.nCopies
    End If
    txtStatus.Text = strStatus
      
    MousePointer = vbDefault
      
ExitPoint:
    Handle_Error Err, gstrObjectName, strFunctionName, "", False, False
End Sub

Private Sub Form_Load()
    On Error GoTo ExitPoint
    
    'Used by the error handling routines
    Const strFunctionName As String = "Form_Load"
       
    ' Indicate window loaded.
    gblnWindowShown = True
    
    ' Set the form caption to the current version.
    frmaxword.Caption = "Active X Word Viewer v" & App.Major & "." & App.Minor & "." & App.Revision
       
    ' Check for errors.
    If gblnError Then
        Unload frmaxword
    Else
        'Setup Busy Mouse pointer
        frmaxword.MousePointer = vbHourglass

        'Make window always on top.
        ' AS 02/11/04 Is this required? Causes problems with IE zorder on exiting Axword.
        'Call SetWindowPos(frmaxword.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)

        If gblnPersistState Then
            LoadState
        End If

        If gblnResizeableFrame Then
            gintHeight = gintCurrentHeight
            gintWidth = gintCurrentWidth
            If gblnPersistState Then
                frmaxword.Left = gintCurrentLeft
                frmaxword.Top = gintCurrentTop
            End If
        Else
            ' Check size params.
            If (gintWidth < 640 Or gintWidth > (Screen.Width / Screen.TwipsPerPixelX)) Then
                gintWidth = (Screen.Width / Screen.TwipsPerPixelX)
                frmaxword.Left = 0
            End If
            If (gintHeight < 480 Or gintHeight > (Screen.Height / Screen.TwipsPerPixelY)) Then
                gintHeight = (Screen.Height / Screen.TwipsPerPixelY)
                frmaxword.Top = 0
            End If
        End If

        ' Resize dialogue.
        frmaxword.Height = gintHeight * Screen.TwipsPerPixelY
        frmaxword.Width = gintWidth * Screen.TwipsPerPixelX
        
        'Initialise
        cmdSave.Enabled = False
        
        'Install windows message hook, technique otherwise known as subclassing.
        'This intercepts all windows messages to this app.
        hMsgHook = SetWindowsHookEx(WH_GETMESSAGE, AddressOf MsgProc, 0&, App.ThreadID)
                
        'Also sub class this window, to prevent resizing.
        Call SubClass(frmaxword.hWnd)
                
        ' Prepare to view or edit.
        
        'Load up the HTML file if available.
        If gblnHTMLGenerated Then
           web1.Navigate gstrFile & ".htm"
        Else
'TW 16/6/2004
'            web1.Navigate "about:blank"
            If gblnPDFGenerated Then
                web1.Navigate gstrFile & ".pdf"
            Else
                web1.Navigate "about:blank"
            End If
'End TW 16/6/2004
        End If

        web1.Visible = True
              
        If gblnReadOnly Then
            frmaxword.cmdSave.Visible = False
            frmaxword.cmdEdit.Visible = False
            gblnShowFindFreeText = False
            If gblnViewAsWord Then
                'Using Word for view mode too - switch to Word.
                frmaxword.Enabled = False
                DoEvents
                frmaxword.cmdEdit_Click
                'Re-enable controls.
                frmaxword.Enabled = True
                frmaxword.cmdExit.Enabled = True
            End If
        Else
            'If we're in Edit mode then go into it straight away.
            frmaxword.cmdEdit.Enabled = False
            frmaxword.cmdExit.Enabled = False
            frmaxword.Enabled = False
            DoEvents
            'Switch to edit mode.
            frmaxword.cmdEdit_Click
            'Re-enable controls.
            frmaxword.Enabled = True
            frmaxword.cmdExit.Enabled = True
        End If
        
        If Not gblnViewAsWord And Not gblnViewAsPDF Then
            gblnShowPrint = False
        End If
        frmaxword.cmdPrint.Visible = gblnShowPrint
        
        If Not gblnShowPrint Then
            gblnShowPrintDialog = False
        End If
        If gblnShowPrintDialog Then
            frmaxword.cmdPrint.Caption = "&Print..."
        End If
              
        ' Setup GUI.
        frmaxword.cmdFindFreeTxt.Visible = gblnShowFindFreeText
        
        frmaxword.CheckTrackedChanges.Visible = gblnShowTrackedChanges
        If gblnShowTrackedChanges And gblnTrackedChanges Then
            frmaxword.CheckTrackedChanges.Enabled = True
            frmaxword.CheckTrackedChanges.Value = False
        Else
            frmaxword.CheckTrackedChanges.Enabled = False
            frmaxword.CheckTrackedChanges.Value = False
        End If

        If Not gblnShowTrackedChanges And Not gblnShowFindFreeText Then
            ' Do not display Options frame if there are no visible controls within it.
            frmaxword.FrameOptions.Visible = False
        End If
        
        'Set startup status.
        txtStatus.Text = mstrStartupStatus
                                
        'Reset Mouse pointer
        MousePointer = vbDefault
         
        DoEvents
    End If
    
ExitPoint:
    Handle_Error Err, gstrObjectName, strFunctionName
 
End Sub

Private Sub LoadState()
    Dim nWindowState As Integer
    nWindowState = GetSetting("Marlborough Stirling", "AxWord", "FormWindowState", vbNormal)
    If nWindowState = vbMinimized Then
        nWindowState = vbNormal
    End If
    frmaxword.WindowState = nWindowState
    gintCurrentHeight = GetSetting("Marlborough Stirling", "AxWord", "FormHeight", 480)
    gintCurrentWidth = GetSetting("Marlborough Stirling", "AxWord", "FormWidth", 640)
    gintCurrentTop = GetSetting("Marlborough Stirling", "AxWord", "FormTop", 0)
    gintCurrentLeft = GetSetting("Marlborough Stirling", "AxWord", "FormLeft", 0)
End Sub

Private Sub SaveState()
    SaveSetting "Marlborough Stirling", "AxWord", "FormHeight", gintCurrentHeight
    SaveSetting "Marlborough Stirling", "AxWord", "FormWidth", gintCurrentWidth
    SaveSetting "Marlborough Stirling", "AxWord", "FormTop", gintCurrentTop
    SaveSetting "Marlborough Stirling", "AxWord", "FormLeft", gintCurrentLeft
    SaveSetting "Marlborough Stirling", "AxWord", "FormWindowState", frmaxword.WindowState
    SavePrintDialogState
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo ExitPoint
    
    Const strFunctionName As String = "Form_QueryUnload"

    ' Unload the form and free objects.
    If (gblnReadOnly = False And gblnError = False) Then
        If gblnEdit = True Then
            If MsgBox("Save changes to document?", vbYesNo, "Save?") = vbYes Then
                Call cmdSave_Click
            Else
                If gblnViewAsWord Then
                    mobjCurrentDocument.Saved = True
                Else
                    'Swap back to viewing the HTML version of the doc in the Web object
                    Dim arrFile() As Byte
    
                    SaveCurrentWordDocument
                    
                    'Close word doc and application.
                    Set mobjCurrentDocument = Nothing
                    web1.Navigate "about:blank"
                    DoEvents
    
                    ' Navigate to the htm version.
                    saveBinBase64AsWordDoc gRequestData.FileContents, gRequestData.nDeliveryType, gstrFile, gblnViewAsWord
                    web1.Navigate gstrFile & ".htm"
                    DoEvents
    
                    ' Now replace the doc with the last version.
                    ' Convert bin.base64 stream back to a byte array.
                    arrFile = ConvertBase64ToBin(gstrObjectName, gRequestData.FileContents, bDecompress:=True)
    
                    ' Save the byte array to disk to produce a word document.
                    Open (gstrFile & gRequestData.strFileExtension) For Binary Access Write Lock Read Write As #1
                    Put #1, , arrFile
                    Close #1
    
                    'Update Global Variables
                    gFileSaved = False
                    gblnEdit = False
                    
                    'Find Free Text is disabled in view mode.
                    frmaxword.cmdFindFreeTxt.Enabled = False
    
                    'Update GUI Buttons
                    cmdEdit.Enabled = True
                    cmdSave.Enabled = False
                End If
            End If
        ElseIf gblnViewAsWord Then
            mobjCurrentDocument.Saved = True
        End If
    ElseIf gblnViewAsWord Then
        mobjCurrentDocument.Saved = True
    End If
    
    MousePointer = vbDefault

    'If theres been some sort of error then we just want to quit
    If Not gblnError Then
        If MsgBox("Are you sure you want to quit?", vbYesNo, "Quit?") = vbNo Then
            Cancel = True
        End If
    End If
    
ExitPoint:
    Handle_Error Err, gstrObjectName, strFunctionName
    
End Sub

Private Sub Form_Resize()
    ' Resize the form.
    On Error GoTo ExitPoint
    
    Const strFunctionName As String = "Form_Resize"
    
    If Not WindowState = vbMinimized Then
        Dim Hgt As Single
        Dim TopPos As Single
    
        If gblnResizeableFrame Then
            gintCurrentWidth = frmaxword.Width / Screen.TwipsPerPixelX
            gintCurrentHeight = frmaxword.Height / Screen.TwipsPerPixelY
        Else
            ' Check size params.
            If (gintWidth < 640 Or gintWidth > (Screen.Width / Screen.TwipsPerPixelX)) Then
                gintWidth = (Screen.Width / Screen.TwipsPerPixelX)
                frmaxword.Left = 0
            End If
    
            If (gintHeight < 480 Or gintHeight > (Screen.Height / Screen.TwipsPerPixelY)) Then
                gintHeight = (Screen.Height / Screen.TwipsPerPixelY)
                frmaxword.Top = 0
            End If
        End If
                    
        ' Position controls.
        Dim nButtons As Integer
        nButtons = 1
        cmdExit.Move 5, ScaleHeight - (cmdExit.Height + 5)
        If cmdEdit.Visible Then
            cmdEdit.Move cmdExit.Left + cmdExit.Width + 5, cmdExit.Top
            nButtons = nButtons + 1
        End If
        If cmdSave.Visible Then
            cmdSave.Move cmdExit.Left + ((cmdExit.Width + 5) * nButtons), cmdExit.Top
            nButtons = nButtons + 1
        End If
        If cmdPrint.Visible Then
            cmdPrint.Move cmdExit.Left + ((cmdExit.Width + 5) * nButtons), cmdExit.Top
            nButtons = nButtons + 1
        End If
        
        FrameOptions.Move 5, ScaleHeight - (cmdExit.Height + 20 + FrameOptions.Height), (ScaleWidth - 10)
        
        cmdFindFreeTxt.Left = (FrameOptions.Width * Screen.TwipsPerPixelX) - (cmdFindFreeTxt.Width + 120)
        
        Hgt = FrameOptions.Top - 10
        If (Hgt < 120) Then
            Hgt = 120
        End If
        If Not FrameOptions.Visible Then
            ' If Options frame is not visible, then we have more room for the main
            ' document window - so increase its height.
            Hgt = Hgt + FrameOptions.Height
        End If
        web1.Move 0, 0, ScaleWidth, Hgt
    
        ' Make sure status is always to the right of the command buttons.
        Dim lngStatusWidth As Long
        Dim cmdButton As CommandButton
        If cmdPrint.Visible Then
            Set cmdButton = cmdPrint
        ElseIf cmdSave.Visible Then
            Set cmdButton = cmdSave
        ElseIf cmdEdit.Visible Then
            Set cmdButton = cmdEdit
        Else
            Set cmdButton = cmdExit
        End If
        lngStatusWidth = (FrameOptions.Left + FrameOptions.Width) - (cmdButton.Left + cmdButton.Width + 10)
        If lngStatusWidth > 0 Then
            txtStatus.Move cmdButton.Left + cmdButton.Width + 10, cmdButton.Top, lngStatusWidth
        End If
    
    End If
    
ExitPoint:
    Handle_Error Err, gstrObjectName, strFunctionName, "", False, False

End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    Const strFunctionName As String = "Form_Unload"
    
    Dim objTempFile As Object
    Dim iPos As Integer
    
    ' Update status.
    frmaxword.txtStatus = "Exiting"
    
    gintCurrentTop = frmaxword.Top
    gintCurrentLeft = frmaxword.Left
    
    If gblnPersistState Then
        SaveState
    End If

    'Check to see if the Windows Hook Routine has been used.
    If (hMsgHook > 0) Then
        Call UnhookWindowsHookEx(hMsgHook)
    End If
    
    Call UnSubClass(Me.hWnd)
              
    If gblnViewAsWord Then
        ' Reset the browser.
        
        SaveCurrentWordDocument
        
        web1.Navigate "about:blank"
    Else
        ' Quit the browser.
        Set mobjCurrentDocument = Nothing
        web1.Stop
    End If
    DoEvents
    
    ' Delete temporary files.
    Call SafeDeleteFile(gstrFile & gRequestData.strFileExtension)
    Call SafeDeleteFile(gstrFile & "_wip" & gRequestData.strFileExtension)
    Call SafeDeleteFile(gstrFile & ".htm")
    Dim objFSO As FileSystemObject
    Set objFSO = New FileSystemObject
    If Not objFSO Is Nothing Then
        ' Delete Word's temp conversion files.
        If objFSO.FolderExists(gstrFile & "_files") Then
            objFSO.DeleteFolder gstrFile & "_files", True
        End If
    End If
    
    ' Clean up.
    Set objFSO = Nothing
End Sub

Private Sub WordToEditMode()
    On Error GoTo ExitPoint
    
    Const strFunctionName As String = "WordToEditMode"
                       
    ' AS 20/09/04 BMIDS850 Only show revisions if ShowTrackedChanges is True.
    If gblnShowTrackedChanges Then
        ' Check if tracked changes are enabled.
        gblnTrackedChanges = mobjCurrentDocument.TrackRevisions
        ' Turn on visible revisions.
        mobjCurrentDocument.ShowRevisions = gblnShowRevisions
    Else
        gblnTrackedChanges = False
        ' Turn off visible revisions.
        mobjCurrentDocument.ShowRevisions = False
    End If
        
    ' Configures the Word OLE object.
    mobjCurrentDocument.Application.CustomizationContext = mobjCurrentDocument.AttachedTemplate
    mobjCurrentDocument.Application.Options.CheckSpellingAsYouType = gblnSpellCheckWhileEditing
    mobjCurrentDocument.ShowGrammaticalErrors = False
    mobjCurrentDocument.ShowSpellingErrors = gblnSpellCheckWhileEditing
    mobjCurrentDocument.Application.DisplayAlerts = wdAlertsNone
    mobjCurrentDocument.Application.NormalTemplate.Saved = True
    mobjCurrentDocument.Application.Options.SavePropertiesPrompt = False
    mobjCurrentDocument.Application.Options.SaveNormalPrompt = False
    mobjCurrentDocument.Application.Browser.Target = wdBrowsePage
    mobjCurrentDocument.ActiveWindow.View.Type = 3 'wdPrintView
    mobjCurrentDocument.ActiveWindow.View.Zoom.PageFit = gintPageFit
    mobjCurrentDocument.ActiveWindow.View.ShowBookmarks = False
    mobjCurrentDocument.SpellingChecked = Not gblnSpellCheckWhileEditing

    If gblnReadOnly = True Or gblnEdit = False Then
        mobjCurrentDocument.Protect wdAllowOnlyComments
    End If
    
    ' Hide forms design.
    ShowFormsDesign gstrObjectName, mobjCurrentDocument, False
    
    ShowCommandBars gblnShowCommandBars, gstrObjectName, mobjCurrentDocument
    
    If Not gblnReadOnly Then
        ' Check whether find free text should be enabled.
        gblnTextSearchEnabled = FindFreeTxt(gstrObjectName, mobjCurrentDocument)
        ' Remove selection
        mobjCurrentDocument.ActiveWindow.Selection.GoTo What:=wdGoToLine, Which:=wdGoToFirst
        
        If gblnTextSearchEnabled And gblnShowFindFreeText Then
            frmaxword.cmdFindFreeTxt.Enabled = True
        Else
            frmaxword.cmdFindFreeTxt.Enabled = False
        End If
    End If
            
    ' Activate document.
    mobjCurrentDocument.Activate
       
ExitPoint:
    Handle_Error Err, gstrObjectName, strFunctionName, "", False, False
    
End Sub


Private Sub web1_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
    On Error GoTo ExitPoint
    
    Debug.Print "In web1_NavigateComplete2: " & URL
    
    Const strFunctionName As String = "web1_NavigateComplete2"
    
    If (LCase(Right$(URL, 4)) = ".doc" Or LCase(Right$(URL, 4)) = ".rtf") Then
        'Make the HTML version disappear and show the Active X Document object.
               
        Set mobjCurrentDocument = pDisp.Document
               
        WordToEditMode
        
        'Sort out the buttons
        cmdSave.Enabled = True
        cmdEdit.Enabled = False
        
        'Update global variables
        'AS 29/01/04 Edit flag now set in cmdEdit_Click if not read only.
        'gblnEdit = True
        MousePointer = vbDefault
        
        ' Update status.
        frmaxword.txtStatus.Text = "Editing document"
       
        DoEvents
    End If
    
ExitPoint:
    Handle_Error Err, gstrObjectName, strFunctionName, "", False, False
End Sub

#End If ' MIN_BUILD
