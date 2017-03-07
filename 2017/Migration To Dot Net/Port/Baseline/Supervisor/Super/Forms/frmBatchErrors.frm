VERSION 5.00
Object = "{24A7A369-0EDB-4924-817E-D81FE8E09890}#1.0#0"; "MSGOCX.ocx"
Begin VB.Form frmBatchErrors 
   Caption         =   "Batch Errors"
   ClientHeight    =   4710
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8910
   Icon            =   "frmBatchErrors.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4710
   ScaleWidth      =   8910
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   7560
      TabIndex        =   2
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Frame fraErrors 
      Caption         =   "Error Information"
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8715
      Begin MSGOCX.MSGListView lvBatchErrors 
         Height          =   3135
         Left            =   240
         TabIndex        =   1
         Top             =   420
         Width           =   8235
         _ExtentX        =   14526
         _ExtentY        =   5530
         Sorted          =   -1  'True
         AllowColumnReorder=   0   'False
      End
   End
End
Attribute VB_Name = "frmBatchErrors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form          : frmBatchErrors
' Description   :
'
' Change history
' Prog      Date        Description
' SA        13/02/02    SYS4037 Created
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' BMids change history
' Prog      Date        Description
' GHun      01/04/2003  BM0425 Aesthetic fixes: improved default column widths and form resizing
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Epsom change history
' Prog      Date        Description
' TW        19/02/2007  EP2_1348 - Error in certain circumstances on MSGListView when AllowColumnReorder set to true
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private m_colKeys As Collection
Private m_clsBatchAudit As BatchAuditTable
Private m_sBatchType  As String

Private Sub cmdOK_Click()
    
    On Error GoTo Failed
         
    Hide

    Exit Sub

Failed:
    g_clsErrorHandling.DisplayError Err.DESCRIPTION
End Sub
Public Sub SetKeys(ByRef colKeys As Collection)
    Set m_colKeys = colKeys
End Sub
Private Sub PopulateScreenControls()
    
    'Setup the column headers on the listview.
    SetLVHeaders
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub SetLVHeaders()
    
    Dim colHeaders As Collection
    Dim lvcolHeaders As listViewAccess
        
    On Error GoTo Failed
    
    'Create a collection to hold column headers.
    Set colHeaders = New Collection
    
    Select Case m_sBatchType
        'Payment Processing
        Case "P"
            'Application Number header.
            lvcolHeaders.nWidth = 17
            lvcolHeaders.sName = "Application No"
            colHeaders.Add lvcolHeaders
    
            'Payment Sequence No header.
            lvcolHeaders.nWidth = 19
            lvcolHeaders.sName = "Payment Seq No"
            colHeaders.Add lvcolHeaders
        'Rate Change
        Case "R"
            'Application Number header.
            lvcolHeaders.nWidth = 16
            lvcolHeaders.sName = "Application No"
            colHeaders.Add lvcolHeaders
    
            'Application Factfind No header.
            lvcolHeaders.nWidth = 14
            lvcolHeaders.sName = "Fact Find No"
            colHeaders.Add lvcolHeaders
            
            'Quote No header.
            lvcolHeaders.nWidth = 11
            lvcolHeaders.sName = "Quote No"
            colHeaders.Add lvcolHeaders
    
            'Mortgage Sub Quote No header.
            lvcolHeaders.nWidth = 23
            lvcolHeaders.sName = "Mortgage Subquote No"
            colHeaders.Add lvcolHeaders
            
            'Loan Component Sequence No
            lvcolHeaders.nWidth = 21
            lvcolHeaders.sName = "Loan Component No"
            colHeaders.Add lvcolHeaders
        'Valuation
        Case "V"
        
            'Application No header.
            lvcolHeaders.nWidth = 16
            lvcolHeaders.sName = "Application No"
            colHeaders.Add lvcolHeaders
    
            'Application Fact Find No header.
            lvcolHeaders.nWidth = 14
            lvcolHeaders.sName = "Fact Find No"
            colHeaders.Add lvcolHeaders
            
            'Instruction SeqNo header.
            lvcolHeaders.nWidth = 20
            lvcolHeaders.sName = "Instruction Seq No"
            colHeaders.Add lvcolHeaders
    
            'Valuer Invoice Amount header.
            lvcolHeaders.nWidth = 23
            lvcolHeaders.sName = "Valuer Invoice Amount"
            colHeaders.Add lvcolHeaders
    End Select
    
    ' All Batch Types will have these columns
    'Error Number header.
    lvcolHeaders.nWidth = 16
    lvcolHeaders.sName = "Error No"
    colHeaders.Add lvcolHeaders
    
    'Error Source header.
    lvcolHeaders.nWidth = 40
    lvcolHeaders.sName = "Error Source"
    colHeaders.Add lvcolHeaders
    
    'Error Description header.
    lvcolHeaders.nWidth = 70
    lvcolHeaders.sName = "Error Description"
    colHeaders.Add lvcolHeaders
    
    'Add the column header to the listview.
    lvBatchErrors.AddHeadings colHeaders
    
Failed:
    Set colHeaders = Nothing
    
    If Err.Number <> 0 Then
        g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
    End If
End Sub
Private Sub Form_Load()
    
    On Error GoTo Failed
        
    Set m_clsBatchAudit = New BatchAuditTable
    
    PopulateScreenControls
    SetScreenFields
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.DisplayError Err.DESCRIPTION
End Sub
Private Sub SetScreenFields(Optional bClear As Boolean = False)
    
    On Error GoTo Failed
    
    ' Tell tableclass what type of batch we are
    m_clsBatchAudit.SetBatchType m_sBatchType
    
    'Get the data passing in BatchNumber and BatchRunNumber
    m_clsBatchAudit.GetBatchAuditTableData m_colKeys(1), m_colKeys(2)
    
    'Populate the listview from the underlying data.
    g_clsFormProcessing.PopulateFromRecordset lvBatchErrors, m_clsBatchAudit, bClear
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

'BM0425
Private Sub Form_Resize()
    Const clBorder      As Long = 240
    Const clMinWidth    As Long = 4400
    Const clMinHeight   As Long = 2800
    
'Ignore resizing errors
On Error Resume Next
    
    If Me.Width < clMinWidth Then
        Me.Width = clMinWidth
    End If
    If Me.Height < clMinHeight Then
        Me.Height = clMinHeight
    End If

    fraErrors.Width = Me.Width - clBorder - 100
    fraErrors.Height = Me.Height - (clBorder * 5)
    lvBatchErrors.Width = fraErrors.Width - (clBorder * 2)
    lvBatchErrors.Height = fraErrors.Height - (clBorder * 3)
    cmdOK.Left = Me.Width - cmdOK.Width - clBorder
    cmdOK.Top = Me.Height - cmdOK.Height - (clBorder * 2) - 60
End Sub
'BM0425 End

Private Sub Form_Unload(Cancel As Integer)
    
    Set m_colKeys = Nothing
    Set m_clsBatchAudit = Nothing
    
End Sub
Public Sub SetBatchType(ByRef sBatchType As String)
    m_sBatchType = sBatchType
End Sub
