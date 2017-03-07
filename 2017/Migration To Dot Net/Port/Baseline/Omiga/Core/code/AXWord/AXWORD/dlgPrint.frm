VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form dlgPrint 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Print"
   ClientHeight    =   2685
   ClientLeft      =   315
   ClientTop       =   3750
   ClientWidth     =   6570
   Icon            =   "dlgPrint.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   6570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrCancel 
      Enabled         =   0   'False
      Left            =   6000
      Top             =   2160
   End
   Begin MSComCtl2.UpDown spinCopies 
      Height          =   285
      Left            =   3136
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   600
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   503
      _Version        =   393216
      Value           =   1
      AutoBuddy       =   -1  'True
      BuddyControl    =   "txtCopies"
      BuddyDispid     =   196613
      OrigLeft        =   3360
      OrigTop         =   600
      OrigRight       =   3600
      OrigBottom      =   855
      Max             =   999
      Min             =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.CheckBox chkDifferentTrayOtherPages 
      Caption         =   "Use &different tray for other pages"
      CausesValidation=   0   'False
      Height          =   255
      Left            =   3720
      TabIndex        =   7
      Top             =   1110
      Width           =   2655
   End
   Begin VB.ComboBox cboOtherPagesTray 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1560
      Width           =   2055
   End
   Begin VB.ComboBox cboFirstPageTray 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1080
      Width           =   2055
   End
   Begin VB.TextBox txtCopies 
      Height          =   285
      Left            =   1560
      TabIndex        =   3
      Top             =   600
      Width           =   1575
   End
   Begin VB.ComboBox cboName 
      CausesValidation=   0   'False
      Height          =   315
      ItemData        =   "dlgPrint.frx":030A
      Left            =   1560
      List            =   "dlgPrint.frx":030C
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   4815
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   2880
      TabIndex        =   11
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1560
      TabIndex        =   10
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label lblOtherPagesTray 
      Caption         =   "&Other pages tray:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1590
      Width           =   1335
   End
   Begin VB.Label lblFirstPageTray 
      Caption         =   "&First page tray:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1110
      Width           =   1335
   End
   Begin VB.Label lblCopies 
      Caption         =   "&Copies:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   610
      Width           =   855
   End
   Begin VB.Label lblName 
      Caption         =   "&Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   130
      Width           =   735
   End
End
Attribute VB_Name = "dlgPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private m_nSelectedPrinter As Integer
Private m_strDeviceName As String
Private m_nCopies As Integer
Private m_nFirstPageTray As Integer
Private m_nOtherPagesTray As Integer
Private m_bOK As Boolean
Private m_bUseDifferentTrayForOtherPages As Boolean
Private m_bShowProgressBar As Boolean
Private m_strDocumentTitle As String

Private Sub Form_Initialize()
    m_nSelectedPrinter = 0
    m_strDeviceName = ""
    m_nFirstPageTray = -1
    m_nOtherPagesTray = -1
    m_nCopies = 1
    m_bOK = False
    m_bUseDifferentTrayForOtherPages = False
    m_bShowProgressBar = False
    m_strDocumentTitle = ""
End Sub


Private Sub Form_Load()
On Error GoTo errorPoint

    Dim bSuccess As Boolean
   
    g_Log.WriteLine "->dlgPrint.Form_Load"
   
    MousePointer = vbHourglass
    
    DoEvents
   
    Dim MessageOp As MessageOpGetPrinters
    If m_bShowProgressBar Then
        Set MessageOp = New MessageOpGetPrinters
        frmMessage.MessageOp = MessageOp
        frmMessage.Caption = "Printing"
        frmMessage.message = "Please wait - getting printer information."
        frmMessage.Show vbModal
        bSuccess = frmMessage.Success
        Set MessageOp = Nothing
    Else
        bSuccess = GetPrinters(gPrinterData)
    End If
    
    WritePrintersToLog gPrinterData
    
    If Not bSuccess Then
        Me.Hide
        MsgBox "Error getting printer information.", vbCritical, "Error"
        tmrCancel.Interval = 10
        tmrCancel.Enabled = True
    Else
    
        If m_strDocumentTitle <> "" Then
            Caption = "Print - " & m_strDocumentTitle
        Else
            Caption = "Print"
        End If
        
        Dim nPrinter As Integer
        m_nSelectedPrinter = 0
        
        g_Log.WriteLine "UBound(gPrinterData) = " & CStr(UBound(gPrinterData))
        
        For nPrinter = 0 To UBound(gPrinterData)
            g_Log.WriteLine "gPrinterData(" & CStr(nPrinter) & ").strDeviceName = " & gPrinterData(nPrinter).strDeviceName
            
            cboName.AddItem gPrinterData(nPrinter).strDeviceName
            If Len(m_strDeviceName) > 0 Then
                ' A printer has already been selected
                If gPrinterData(nPrinter).strDeviceName = m_strDeviceName Then
                    m_nSelectedPrinter = nPrinter
                End If
            Else
                ' Select the default printer.
                If gPrinterData(nPrinter).bDefault Then
                    m_nSelectedPrinter = nPrinter
                End If
            End If
        Next
        
        If m_nFirstPageTray = -1 Then
            ' Not set in caller, so use default bin.
            m_nFirstPageTray = wdPrinterDefaultBin
        End If
        If m_nOtherPagesTray = -1 Then
            ' Not set in caller, so use default bin.
            m_nOtherPagesTray = wdPrinterDefaultBin
        End If
        
        If cboName.ListCount > m_nSelectedPrinter Then
            ' Will call cboName_Click as side effect, which populates tray combo boxes.
            cboName.ListIndex = m_nSelectedPrinter
        End If
                
        txtCopies.Text = m_nCopies
       
        If m_bUseDifferentTrayForOtherPages Then
            chkDifferentTrayOtherPages.Value = 1
        Else
            chkDifferentTrayOtherPages.Value = 0
        End If
        chkDifferentTrayOtherPages_Click
              
        Call SetWindowPos(hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
                      
    End If

ExitPoint:

    MousePointer = vbDefault

    g_Log.WriteLine "<-dlgPrint.Form_Load"
    
    Exit Sub
    
errorPoint:

    g_Log.WriteLine "Error in dlgPrint.Form_Load: " & Err.Description
    
    GoTo ExitPoint
    
End Sub

Private Sub cboName_Click()
    m_nSelectedPrinter = cboName.ListIndex
    
    Dim nBin As Integer
    Dim nSelectedFirstPageTray As Integer
    Dim nSelectedOtherPagesTray As Integer
    
    m_nFirstPageTray = -1
    m_nOtherPagesTray = -1
    
    nSelectedFirstPageTray = -1
    nSelectedOtherPagesTray = -1
   
    cboFirstPageTray.Clear
    cboOtherPagesTray.Clear
    Debug.Print "Device name: " & gPrinterData(m_nSelectedPrinter).strDeviceName
    Debug.Print "Default bin: " & gPrinterData(m_nSelectedPrinter).nDefaultBin
    If gPrinterData(m_nSelectedPrinter).nBins > 0 Then
        For nBin = 0 To UBound(gPrinterData(m_nSelectedPrinter).strBinNames)
            
            cboFirstPageTray.AddItem gPrinterData(m_nSelectedPrinter).strBinNames(nBin), nBin
            cboOtherPagesTray.AddItem gPrinterData(m_nSelectedPrinter).strBinNames(nBin), nBin
            cboFirstPageTray.ItemData(nBin) = gPrinterData(m_nSelectedPrinter).nBinNumbers(nBin)
            cboOtherPagesTray.ItemData(nBin) = gPrinterData(m_nSelectedPrinter).nBinNumbers(nBin)
            
            Debug.Print "Bin index: " & nBin
            Debug.Print "Bin number: " & gPrinterData(m_nSelectedPrinter).nBinNumbers(nBin)
            If m_nFirstPageTray > -1 Then
                If gPrinterData(m_nSelectedPrinter).nBinNumbers(nBin) = m_nFirstPageTray Then
                    nSelectedFirstPageTray = nBin
                End If
            Else
                If gPrinterData(m_nSelectedPrinter).nBinNumbers(nBin) = gPrinterData(m_nSelectedPrinter).nDefaultBin Then
                    nSelectedFirstPageTray = nBin
                End If
            End If
            
            If m_nOtherPagesTray > -1 Then
                If gPrinterData(m_nSelectedPrinter).nBinNumbers(nBin) = m_nOtherPagesTray Then
                    nSelectedOtherPagesTray = nBin
                End If
            Else
                If gPrinterData(m_nSelectedPrinter).nBinNumbers(nBin) = gPrinterData(m_nSelectedPrinter).nDefaultBin Then
                    nSelectedOtherPagesTray = nBin
                End If
            End If
            
        Next
    End If
    
    cboFirstPageTray.Enabled = IIf(cboFirstPageTray.ListCount > 0, True, False)
    lblFirstPageTray.Enabled = IIf(cboFirstPageTray.ListCount > 0, True, False)
    chkDifferentTrayOtherPages.Enabled = IIf(cboOtherPagesTray.ListCount > 0, True, False)
    chkDifferentTrayOtherPages.Value = 0
    chkDifferentTrayOtherPages_Click
    
    If nSelectedFirstPageTray > -1 And nSelectedFirstPageTray < cboFirstPageTray.ListCount Then
        cboFirstPageTray.ListIndex = nSelectedFirstPageTray
        m_nFirstPageTray = cboFirstPageTray.ItemData(nSelectedFirstPageTray)
    Else
        If cboFirstPageTray.ListCount > 0 Then
            ' First item is always Default tray.
            cboFirstPageTray.ListIndex = 0
        End If
        m_nFirstPageTray = gPrinterData(m_nSelectedPrinter).nDefaultBin
    End If

    If nSelectedOtherPagesTray > -1 And nSelectedOtherPagesTray < cboOtherPagesTray.ListCount Then
        cboOtherPagesTray.ListIndex = nSelectedOtherPagesTray
        m_nOtherPagesTray = cboOtherPagesTray.ItemData(nSelectedOtherPagesTray)
    Else
        If cboOtherPagesTray.ListCount > 0 Then
            ' First item is always Default tray.
            cboOtherPagesTray.ListIndex = 0
        End If
        m_nOtherPagesTray = gPrinterData(m_nSelectedPrinter).nDefaultBin
    End If

End Sub

Private Sub OKButton_Click()
    m_strDeviceName = gPrinterData(m_nSelectedPrinter).strDeviceName

    cboFirstPageTray_Click
    cboOtherPagesTray_Click
    
    m_nCopies = txtCopies.Text
    
    m_bOK = True
    
    Unload Me
End Sub

Private Sub cboFirstPageTray_Click()
    Dim nSelectedFirstPageTray As Integer
    nSelectedFirstPageTray = cboFirstPageTray.ListIndex
    If nSelectedFirstPageTray > -1 And nSelectedFirstPageTray < cboFirstPageTray.ListCount Then
        m_nFirstPageTray = cboFirstPageTray.ItemData(nSelectedFirstPageTray)
        
        If chkDifferentTrayOtherPages.Value = 0 Then
            ' Other pages are using the same tray as first page, so keep other pages tray combo box
            ' in sync.
            cboOtherPagesTray.ListIndex = nSelectedFirstPageTray
        End If
    End If
End Sub

Private Sub cboOtherPagesTray_Click()
    Dim nSelectedOtherPagesTray As Integer
    nSelectedOtherPagesTray = cboOtherPagesTray.ListIndex
    If nSelectedOtherPagesTray > -1 And nSelectedOtherPagesTray < cboOtherPagesTray.ListCount Then
        m_nOtherPagesTray = cboOtherPagesTray.ItemData(nSelectedOtherPagesTray)
    End If
End Sub

Private Sub CancelButton_Click()
    m_bOK = False
    Unload Me
End Sub

Public Property Get DeviceName() As String
    DeviceName = m_strDeviceName
End Property

Public Property Let DeviceName(strDeviceName As String)
    m_strDeviceName = strDeviceName
End Property

Public Property Get Copies() As Integer
    Copies = m_nCopies
End Property

Public Property Let Copies(nCopies As Integer)
    m_nCopies = nCopies
End Property

Public Property Get FirstPageTray() As Integer
    FirstPageTray = m_nFirstPageTray
End Property

Public Property Let FirstPageTray(nFirstPageTray As Integer)
    m_nFirstPageTray = nFirstPageTray
End Property

Public Property Get OtherPagesTray() As Integer
    OtherPagesTray = m_nOtherPagesTray
End Property

Public Property Let UseDifferentTrayForOtherPages(bUseDifferentTrayForOtherPages As Boolean)
    m_bUseDifferentTrayForOtherPages = bUseDifferentTrayForOtherPages
End Property

Public Property Get UseDifferentTrayForOtherPages() As Boolean
    UseDifferentTrayForOtherPages = m_bUseDifferentTrayForOtherPages
End Property

Public Property Let OtherPagesTray(nOtherPagesTray As Integer)
    m_nOtherPagesTray = nOtherPagesTray
End Property

Public Property Get OK() As Boolean
    OK = m_bOK
End Property

Public Property Get ShowProgressBar() As Variant
    ShowProgressBar = m_bShowProgressBar
End Property

Public Property Let ShowProgressBar(ByVal vNewValue As Variant)
    m_bShowProgressBar = vNewValue
End Property

Public Property Get DocumentTitle() As String
    DocumentTitle = m_strDocumentTitle
End Property

Public Property Let DocumentTitle(strDocumentTitle As String)
    m_strDocumentTitle = strDocumentTitle
End Property


Private Sub chkDifferentTrayOtherPages_Click()
    If chkDifferentTrayOtherPages.Value = 1 And cboOtherPagesTray.ListCount > 0 Then
        lblOtherPagesTray.Enabled = True
        cboOtherPagesTray.Enabled = True
        m_bUseDifferentTrayForOtherPages = True
    Else
        lblOtherPagesTray.Enabled = False
        cboOtherPagesTray.Enabled = False
        cboOtherPagesTray.ListIndex = cboFirstPageTray.ListIndex
        m_bUseDifferentTrayForOtherPages = False
    End If
End Sub

Private Sub tmrCancel_Timer()
    tmrCancel.Enabled = False
    CancelButton_Click
End Sub
