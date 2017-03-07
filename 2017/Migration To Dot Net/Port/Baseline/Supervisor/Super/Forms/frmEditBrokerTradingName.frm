VERSION 5.00
Object = "{24A7A369-0EDB-4924-817E-D81FE8E09890}#1.0#0"; "MSGOCX.ocx"
Begin VB.Form frmEditBrokerTradingName 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add/Edit Trading Name"
   ClientHeight    =   2280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9165
   Icon            =   "frmEditBrokerTradingName.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   9165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Details"
      Height          =   1575
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   8895
      Begin MSGOCX.MSGEditBox txtEffectiveDate 
         Height          =   285
         Left            =   1560
         TabIndex        =   2
         Top             =   1080
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         Mask            =   "##/##/####"
         TextType        =   1
         FontSize        =   8.25
         FontName        =   "MS Sans Serif"
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
      Begin MSGOCX.MSGEditBox txtAlternativeName 
         Height          =   285
         Left            =   1560
         TabIndex        =   0
         Top             =   360
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   503
         TextType        =   4
         PromptInclude   =   0   'False
         FontSize        =   8.25
         FontName        =   "MS Sans Serif"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   0
         EndProperty
      End
      Begin MSGOCX.MSGComboBox cboFirmNameType 
         Height          =   315
         Left            =   1560
         TabIndex        =   1
         Top             =   720
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ListIndex       =   -1
         Text            =   ""
      End
      Begin MSGOCX.MSGEditBox txtPrincipalFirmID 
         Height          =   285
         Left            =   7440
         TabIndex        =   6
         Top             =   720
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         TextType        =   4
         PromptInclude   =   0   'False
         FontSize        =   8.25
         FontName        =   "MS Sans Serif"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   0
         EndProperty
      End
      Begin MSGOCX.MSGEditBox txtARFirmID 
         Height          =   285
         Left            =   7440
         TabIndex        =   7
         Top             =   1080
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         TextType        =   4
         PromptInclude   =   0   'False
         FontSize        =   8.25
         FontName        =   "MS Sans Serif"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   0
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Firm Name Type"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   10
         Top             =   720
         Width           =   1155
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Trading Name"
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Effective Date"
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   8
         Top             =   1080
         Width           =   1020
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   6360
      TabIndex        =   3
      Top             =   1800
      Width           =   1275
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   7740
      TabIndex        =   4
      Top             =   1800
      Width           =   1275
   End
End
Attribute VB_Name = "frmEditBrokerTradingName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Change history
' Prog      Date        Description
' AW        03/10/07    OMIGA00003301   Make Border Style Fixed Dialog
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
' Private data
Private m_clsFirmTradingNameTable As FirmTradingNameTable

Private m_bIsEdit As Boolean
Private m_ReturnCode As MSGReturnCode
Private m_clsTableAccess As TableAccess
Private m_colKeys As Collection

Private strFirmTradingNameSeqNo As String

Public Function GetReturnCode() As MSGReturnCode
    GetReturnCode = m_ReturnCode
End Function

Private Sub GetNewFirmTradingNameSeqNo()
Dim conn As ADODB.Connection
Dim cmd As ADODB.Command

    On Error GoTo Failed

    Set conn = New ADODB.Connection
    Set cmd = New ADODB.Command

    With conn
        .ConnectionString = g_clsDataAccess.GetActiveConnection
        .Open
    End With

    With cmd
        .ActiveConnection = conn
        .CommandType = adCmdStoredProc
        .CommandText = "USP_GETNEXTSEQUENCENUMBER"
        .Parameters.Append .CreateParameter("SequenceName", adVarChar, adParamInput, 50, "FirmTradingNameSeqNo")
        .Parameters.Append .CreateParameter("NextNumber", adVarChar, adParamOutput, 12)
        .Execute , , adExecuteNoRecords
        strFirmTradingNameSeqNo = .Parameters("NextNumber").Value
    End With
    
    'Close the database connection
    conn.Close
    Set cmd.ActiveConnection = Nothing
    Set cmd = Nothing
    Set conn = Nothing

    If Len(strFirmTradingNameSeqNo) = 0 Then
        MsgBox "Can't allocate an identity number for this record", vbCritical
    End If
Failed:
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   DoOKProcessing
' Description   :   Common function used when the user presses ok
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function DoOKProcessing() As Boolean
    Dim bRet As Boolean

    bRet = g_clsFormProcessing.DoMandatoryProcessing(Me)
    
    If bRet = True Then
        SaveScreenData
        SaveChangeRequest
    End If

    DoOKProcessing = bRet
    
    Exit Function
Failed:
    DoOKProcessing = False
    g_clsErrorHandling.DisplayError
End Function

Private Sub SaveChangeRequest()
    On Error GoTo Failed
    Dim sProductNumber As String
    Dim colMatchValues As Collection
    Dim clsTableAccess As TableAccess
    
    Set colMatchValues = New Collection
    colMatchValues.Add strFirmTradingNameSeqNo
    
    Set clsTableAccess = m_clsFirmTradingNameTable
    clsTableAccess.SetKeyMatchValues colMatchValues
    g_clsHandleUpdates.SaveChangeRequest m_clsFirmTradingNameTable
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SetEditState
' Description   :   Specific code when the user is editing FirmTradingName
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetEditState()
    
Dim colDataSet As New Collection
Dim strFirmNameType As String

    colDataSet.Add m_colKeys.Item(1)
    TableAccess(m_clsFirmTradingNameTable).SetKeyMatchValues colDataSet
    TableAccess(m_clsFirmTradingNameTable).GetTableData

    txtPrincipalFirmID.Text = m_clsFirmTradingNameTable.GetPrincipalFirmID()
    txtARFirmID.Text = m_clsFirmTradingNameTable.GetARFirmID()
    
    strFirmTradingNameSeqNo = m_clsFirmTradingNameTable.GetFirmTradingNameSeqNo()
    txtAlternativeName.Text = m_clsFirmTradingNameTable.GetAlternativeName()
    Select Case m_clsFirmTradingNameTable.GetFirmNameType
        Case "1" To "3"
            cboFirmNameType.ListIndex = Val(m_clsFirmTradingNameTable.GetFirmNameType) - 1
        Case Else
            cboFirmNameType.ListIndex = -1
    End Select
    txtEffectiveDate.Text = m_clsFirmTradingNameTable.GetEffectiveDate


End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SaveScreenData
' Description   :   Saves all screen data to the database
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SaveScreenData()
Dim vTmp As Variant
Dim clsTableAccess As TableAccess
Dim vNextNumber As Variant
    
    On Error GoTo Failed
    Set clsTableAccess = m_clsFirmTradingNameTable
    
    m_clsFirmTradingNameTable.SetPrincipalfirmID ""
    m_clsFirmTradingNameTable.SetARFirmID txtARFirmID.Text
    m_clsFirmTradingNameTable.SetFirmTradingNameSeqNo strFirmTradingNameSeqNo
    m_clsFirmTradingNameTable.SetAlternativeName txtAlternativeName.Text
    m_clsFirmTradingNameTable.SetFirmNameType cboFirmNameType.ListIndex + 1
    m_clsFirmTradingNameTable.SetEffectiveDate txtEffectiveDate.Text
    m_clsFirmTradingNameTable.SetLastupdateddate Now
    m_clsFirmTradingNameTable.SetLastUpdatedBy g_sSupervisorUser
      
    clsTableAccess.Update
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


Public Sub SetKeys(colKeys As Collection)
    Set m_colKeys = colKeys
End Sub


Public Sub SetIsEdit(Optional bIsEdit As Boolean = True)
    m_bIsEdit = bIsEdit
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SetAddState
' Description   :   Specific code when the user is adding a new FirmTradingName
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetAddState()
    On Error GoTo Failed
    
    Call GetNewFirmTradingNameSeqNo
    'Create a key's collection to set on all child objects.
    Set m_colKeys = New Collection
           
    'Add this generated GUID into the keys collection.
    m_colKeys.Add strFirmTradingNameSeqNo
    
    'Create empty records.
    g_clsFormProcessing.CreateNewRecord m_clsFirmTradingNameTable
           
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub cmdCancel_Click()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cmdCancel_Click
' Description   : Called when the user presses the Cancel button.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Hide
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub

Private Sub cmdOK_Click()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cmdOK_Click
' Description   : Called when the user presses the OK button.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim bRet As Boolean
    On Error GoTo Failed
    
    bRet = DoOKProcessing()

    If bRet = True Then
        
        SetReturnCode
        Hide
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError

End Sub

Private Sub Form_Activate()
    On Error GoTo Failed
    
    ' Initialise Form
    SetReturnCode MSGFailure
    
    Set m_clsTableAccess = New FirmTradingNameTable
    
    Set m_clsFirmTradingNameTable = New FirmTradingNameTable
    
    If m_bIsEdit = True Then
        SetEditState
    Else
        SetAddState
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION


End Sub

Private Sub SetReturnCode(Optional enumReturn As MSGReturnCode = MSGSuccess)
    m_ReturnCode = enumReturn
End Sub


Private Sub Form_Load()
    
    cboFirmNameType.AddItem "Registered Name"
    cboFirmNameType.AddItem "Trading Name"
    cboFirmNameType.AddItem "Common Name"
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set m_clsFirmTradingNameTable = Nothing
End Sub

Private Sub txtAlternativeName_Validate(Cancel As Boolean)
    Cancel = Not txtAlternativeName.ValidateData()
End Sub


Private Sub txtEffectiveDate_Validate(Cancel As Boolean)
    Cancel = Not txtEffectiveDate.ValidateData()
End Sub



