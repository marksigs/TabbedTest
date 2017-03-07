VERSION 5.00
Object = "{24A7A369-0EDB-4924-817E-D81FE8E09890}#1.0#0"; "MSGOCX.ocx"
Begin VB.Form frmProductErrors 
   Caption         =   "Mortgage Product Errors"
   ClientHeight    =   3660
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8565
   Icon            =   "frmProductErrors.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3660
   ScaleWidth      =   8565
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   5520
      TabIndex        =   2
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   6840
      TabIndex        =   1
      Top             =   3000
      Width           =   1215
   End
   Begin MSGOCX.MSGListView lvProductErrors 
      Height          =   1935
      Left            =   420
      TabIndex        =   0
      Top             =   780
      Width           =   7635
      _ExtentX        =   13467
      _ExtentY        =   3413
      Sorted          =   -1  'True
      AllowColumnReorder=   0   'False
   End
   Begin VB.Label lblStartDate 
      Height          =   255
      Left            =   1980
      TabIndex        =   6
      Top             =   480
      Width           =   2475
   End
   Begin VB.Label lblProductCode 
      Height          =   255
      Left            =   1980
      TabIndex        =   5
      Top             =   120
      Width           =   3615
   End
   Begin VB.Label Label2 
      Caption         =   "Start Date"
      Height          =   255
      Left            =   540
      TabIndex        =   4
      Top             =   480
      Width           =   1275
   End
   Begin VB.Label Label1 
      Caption         =   "Product Code"
      Height          =   255
      Left            =   540
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmProductErrors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Epsom Change history
' Prog      Date        Description
' TW        19/02/2007  EP2_1348 - Error in certain circumstances on MSGListView when AllowColumnReorder set to true
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit
Private m_colMatchValues As Collection
Private m_ReturnCode As MSGReturnCode
Private Const PRODUCT_CODE_KEY As Integer = 1
Private Const START_DATE_KEY As Integer = 2
Public Sub SetKeyMatchValues(colMatchValues As Collection)
    On Error GoTo Failed
    
    If Not colMatchValues Is Nothing Then
        Set m_colMatchValues = colMatchValues
    Else
        g_clsErrorHandling.RaiseError errGeneralError, "Product Errors - product keys are empty"
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub cmdCancel_Click()
    On Error GoTo Failed
    
    Hide
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
Private Sub cmdOK_Click()
    On Error GoTo Failed
    SetReturnCode
    Hide
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub

Private Sub Form_Load()
    On Error GoTo Failed
    SetReturnCode MSGFailure
    
    SetProducInfo
    SetHeaders
    ShowErrors

    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
    g_clsFormProcessing.CancelForm Me
End Sub
Private Sub SetHeaders()
    On Error GoTo Failed
    Dim headers As New Collection
    Dim lvHeaders As listViewAccess
    
    lvHeaders.nWidth = 150
    lvHeaders.sName = "Reason"
    headers.Add lvHeaders
    
    lvProductErrors.AddHeadings headers
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub ShowErrors()
    On Error GoTo Failed
    Dim clsProductErrors As InvalidProductTable
    
    If Not m_colMatchValues Is Nothing Then
        Set clsProductErrors = New InvalidProductTable
        TableAccess(clsProductErrors).SetKeyMatchValues m_colMatchValues
        clsProductErrors.GetErrors
        
        g_clsFormProcessing.PopulateFromRecordset lvProductErrors, clsProductErrors
        If lvProductErrors.ListItems.Count > 0 Then
            Set lvProductErrors.SelectedItem = lvProductErrors.ListItems.Item(1)
        End If
    Else
        g_clsErrorHandling.RaiseError errKeysEmpty, "Product Errors "
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
Private Sub SetProducInfo()
    On Error GoTo Failed
    Dim sStartDate As String
    Dim sProductCode As String
    
    sProductCode = m_colMatchValues(PRODUCT_CODE_KEY)
    sStartDate = m_colMatchValues(START_DATE_KEY)
    
    lblProductCode.Caption = sProductCode
    lblStartDate.Caption = sStartDate
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
