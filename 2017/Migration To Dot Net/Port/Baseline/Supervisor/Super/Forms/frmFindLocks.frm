VERSION 5.00
Object = "{61B85B2D-AE68-4B98-B8E9-B759DBAF18E8}#1.1#0"; "MSGOCX.ocx"
Begin VB.Form frmFindLocks 
   Caption         =   "Find Lock"
   ClientHeight    =   2115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6135
   Icon            =   "frmFindLocks.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   6135
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Default         =   -1  'True
      Height          =   375
      Left            =   4830
      TabIndex        =   4
      Top             =   810
      Width           =   1215
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
      Height          =   375
      Left            =   4830
      TabIndex        =   3
      Top             =   270
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1650
      Width           =   1215
   End
   Begin MSGOCX.MSGEditBox txtSearch 
      Height          =   315
      Index           =   0
      Left            =   1920
      TabIndex        =   0
      Top             =   120
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   556
      TextType        =   4
      FontSize        =   8.25
      FontName        =   "MS Sans Serif"
      BackColor       =   16777215
      BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
         Type            =   0
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   0
      EndProperty
      MaxLength       =   12
   End
   Begin MSGOCX.MSGEditBox txtSearch 
      Height          =   315
      Index           =   1
      Left            =   1920
      TabIndex        =   1
      Top             =   570
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   556
      TextType        =   4
      FontSize        =   8.25
      FontName        =   "MS Sans Serif"
      BackColor       =   16777215
      BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
         Type            =   0
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   0
      EndProperty
      MaxLength       =   10
   End
   Begin MSGOCX.MSGEditBox txtSearch 
      Height          =   315
      Index           =   2
      Left            =   1920
      TabIndex        =   2
      Top             =   1050
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   556
      TextType        =   4
      FontSize        =   8.25
      FontName        =   "MS Sans Serif"
      BackColor       =   16777215
      BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
         Type            =   0
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   0
      EndProperty
      MaxLength       =   10
   End
   Begin VB.Label Label2 
      Caption         =   "User ID"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1110
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Unit ID"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   630
      Width           =   1695
   End
   Begin VB.Label lblAppNumber 
      Caption         =   "Application Number"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   180
      Width           =   1695
   End
End
Attribute VB_Name = "frmFindLocks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form          : frmFindLocks
' Description   : Finds all locks matching the search criteria
'
' Change history
' Prog      Date        Description
' DJP       27/02/03    BM0282 Added new search criteria
' IK        07/03/03    BM0314 Added remove (DMS) Document Locks
' BS        28/04/03    BM0282 Clear Listview and allow search on UserID
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

' Private data
Private m_enumLockType As LockType
Private m_clsTableAccess As TableAccess
Private m_ReturnCode As MSGReturnCode
Private m_colKeys As Collection
'BS BM0282 28/04/03
Private m_lvResults As MSGListView

' Constants
Private Const SEARCH_APPNO = 0
Private Const SEARCH_UNITID = 1
Private Const SEARCH_USERID = 2


Public Sub SetTableAccess(clsTableAccess As TableAccess)
    Set m_clsTableAccess = clsTableAccess
End Sub
Public Sub SetKeys(colKeys As Collection)
    Set m_colKeys = colKeys
End Sub
Public Sub SetLockType(enumType As LockType)
    m_enumLockType = enumType
End Sub

Private Sub cmdCancel_Click()
    Hide
End Sub
Private Sub cmdOK_Click()
    On Error GoTo Failed
    
    Hide
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
Private Sub cmdSearch_Click()
    On Error GoTo Failed
    
    DoSearch
    SetReturnCode
    Hide
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub

Private Sub cmdClear_Click()
    On Error GoTo Failed
    
    'BS BM0282 28/04/03
    'Clear all items from the listview.
    m_lvResults.ListItems.Clear
    
    g_clsFormProcessing.ClearScreenFields Me
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub

Private Sub Form_Load()
    On Error GoTo Failed
    SetReturnCode MSGFailure
    
    'BS BM0282 28/04/03
    Set m_lvResults = frmMain.lvListView
    
    Select Case m_enumLockType
    
    Case Application
        'Set m_clsTableAccess = New ApplicationLockTable
        Caption = "Find Application Lock"
        lblAppNumber.Caption = "Application Number"
    Case Customer
        'Set m_clsTableAccess = New CustomerLockTable
        Caption = "Find Customer Lock"
        lblAppNumber.Caption = "Customer Number"
        
    ' ik_BM0314
    Case Document
        Caption = "Find Document Lock"
        lblAppNumber.Caption = "Application Number"
        Label1.Visible = False
        Label2.Visible = False
        txtSearch(1).Visible = False
        txtSearch(2).Visible = False
    
    End Select
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
Private Sub DoSearch()
    On Error GoTo Failed
    Dim sError As String
    Dim sNumber As String
    Dim sUserID As String
    Dim sUnitID As String
    Dim bFound As Boolean
    
    ' Get the application/customer number
    sNumber = txtSearch(SEARCH_APPNO).Text
        
    If Len(sNumber) > 0 Then
        sNumber = g_clsSQLAssistSP.FormatWildcardedString(sNumber, bFound)
    End If
        
    If m_enumLockType = Document Then
            
        If Len(sNumber) = 0 Then
            sError = "Application Number must be entered"
            txtSearch(SEARCH_APPNO).SetFocus
            g_clsErrorHandling.RaiseError errGeneralError, sError
        End If
    
    Else
        
        'BS BM0282 28/04/03
        sUserID = txtSearch(SEARCH_USERID).Text
        
        If Len(sUserID) > 0 Then
            sUserID = g_clsSQLAssistSP.FormatWildcardedString(sUserID, bFound)
        End If
            
        sUnitID = txtSearch(SEARCH_UNITID).Text
            
        If Len(sUnitID) > 0 Then
            sUnitID = g_clsSQLAssistSP.FormatWildcardedString(sUnitID, bFound)
        End If
            
        If Len(sNumber) = 0 And Len(sUserID) = 0 And Len(sUnitID) = 0 Then
            sError = "At least one of "
            If m_enumLockType = Application Then
                sError = sError & " Application Lock"
            Else
                sError = sError & " Customer Lock"
            End If
            
            sError = sError & ", User ID and Unit ID must be entered"
            txtSearch(SEARCH_APPNO).SetFocus
            g_clsErrorHandling.RaiseError errGeneralError, sError
        End If
        
    End If
        
    If Not m_clsTableAccess Is Nothing Then
        Select Case m_enumLockType
        Case Application
            Dim clsAppLocks As ApplicationLockTable
            Set clsAppLocks = m_clsTableAccess
            clsAppLocks.GetApplicationLocks , sNumber, sUserID, sUnitID
        
        Case Customer
            Dim clsCustomerLocks As CustomerLockTable
            Set clsCustomerLocks = m_clsTableAccess
            clsCustomerLocks.GetCustomerLocks , sNumber, sUserID, sUnitID
        
        ' ik_bm0314
        Case Document
            Dim clsDocumentLocks As DocumentLock
            Set clsDocumentLocks = m_clsTableAccess
            ' ik_wip_20030310
            ' clsDocumentLocks.GetDocumentLocks sNumber, sUserID, sUnitID
            clsDocumentLocks.GetDocumentLocks sNumber
        
        End Select
        
        m_clsTableAccess.ValidateData
        SetReturnCode
    Else
        g_clsErrorHandling.RaiseError errGeneralError, "FindLocks: Table class is empty"
    End If
    
    If m_clsTableAccess.RecordCount = 0 Then
        g_clsErrorHandling.RaiseError errGeneralError, "No records found for your search criteria"
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub SetReturnCode(Optional enumReturn As MSGReturnCode = MSGSuccess)
    m_ReturnCode = enumReturn
End Sub
Public Function GetReturnCode() As MSGReturnCode
    GetReturnCode = m_ReturnCode
End Function

