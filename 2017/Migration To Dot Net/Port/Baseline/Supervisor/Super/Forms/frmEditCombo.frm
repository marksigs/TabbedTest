VERSION 5.00
Object = "{5F540CC8-EA22-4F95-9EFE-BDB4E09F976D}#1.0#0"; "MSGOCX.ocx"
Begin VB.Form frmEditCombo 
   Caption         =   "Add/Edit Combo Box Group"
   ClientHeight    =   7950
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5460
   Icon            =   "frmEditCombo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   5460
   StartUpPosition =   1  'CenterOwner
   Begin MSGOCX.MSGTextMulti txtComboNotes 
      Height          =   855
      Left            =   600
      TabIndex        =   10
      Top             =   6480
      Width           =   4755
      _ExtentX        =   8387
      _ExtentY        =   1508
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   ""
      MaxLength       =   255
   End
   Begin VB.CommandButton cmdAnother 
      Caption         =   "&Another"
      Height          =   375
      Left            =   4140
      TabIndex        =   5
      Top             =   7440
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Validation"
      Height          =   2895
      Left            =   60
      TabIndex        =   8
      Top             =   3480
      Width           =   5295
      Begin VB.CommandButton cmdValidationEdit 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   1680
         TabIndex        =   12
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdValidationAdd 
         Caption         =   "&Add"
         Height          =   375
         Left            =   600
         TabIndex        =   11
         Top             =   240
         Width           =   975
      End
      Begin MSGOCX.MSGDataGrid dgValidation 
         Height          =   2475
         Left            =   360
         TabIndex        =   2
         Top             =   240
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   4366
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AllowAdd        =   0   'False
         AllowDelete     =   0   'False
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Values"
      Height          =   2775
      Left            =   60
      TabIndex        =   7
      Top             =   600
      Width           =   5295
      Begin VB.CommandButton cmdValueEdit 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   1680
         TabIndex        =   14
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdValueAdd 
         Caption         =   "&Add"
         Height          =   375
         Left            =   600
         TabIndex        =   13
         Top             =   240
         Width           =   975
      End
      Begin MSGOCX.MSGDataGrid dgValues 
         Height          =   2415
         Left            =   360
         TabIndex        =   1
         Top             =   180
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   4260
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AllowAdd        =   0   'False
         AllowDelete     =   0   'False
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   1500
      TabIndex        =   3
      Top             =   7440
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   2820
      TabIndex        =   4
      Top             =   7440
      Width           =   1215
   End
   Begin MSGOCX.MSGEditBox txtComboBox 
      Height          =   315
      Index           =   0
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   4035
      _ExtentX        =   7117
      _ExtentY        =   556
      TextType        =   4
      PromptInclude   =   0   'False
      FontSize        =   8.25
      FontName        =   "MS Sans Serif"
      Mandatory       =   -1  'True
      BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
         Type            =   0
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   0
      EndProperty
      MaxLength       =   30
   End
   Begin VB.Label Label2 
      Caption         =   "Notes"
      Height          =   255
      Left            =   60
      TabIndex        =   9
      Top             =   6540
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Group Name"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   180
      Width           =   1095
   End
End
Attribute VB_Name = "frmEditCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form          :   frmEditCombo
' Description   :   Form which allows the user add and edit combo values
' Change history
' Prog      Date        Description
' DJP       04/08/01    DJP don't log fact that Omiga 3 Value ID doesn't exist
' STB       06/12/01    SYS1942 - Another button commits current transaction.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'BMIDS Changes
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'CL     08/10/02    BMIDS00565    Modification to DoesChildRecordExist
'                           To handle non exitant rows in combo's
'MV     07/01/03    BM0085  Amended to support Creating audit Records
'MV     15/01/03    BM0085  Amended to support Creating audit Records
'BS     01/05/03    BM0498  Remove leading blank from change description
'JD     30/03/2005  BMIDS982 Save the group note text when changed
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Private Const GROUP_NAME_VAL As Integer = 0

' The Value DataGrid
Private Const VALUE_GROUP_COL = 0
Private Const VALUE_ID_COL = 1
Private Const VALUE_NAME_COL = 2
Private Const VALUE_TITLE = "Value ID"
' The Validation DataGrid
Private Const VALIDATION_NAME_COL = 0
Private Const VALIDATION_VALUE_COL = 1
Private Const VALIDATION_TYPE_COL = 2

' Private data
Private m_bIsEdit As Boolean

' Record sets for all three tables being used
Private m_clsComboValues As ComboValueTable
Private m_clsComboGroup As ComboValueGroupTable
Private m_clsComboUtils As New ComboUtils
Private m_clsComboValidation  As New ComboValidationTable

Private m_sGroupNameOrig As String
Private m_sGroupNoteOrig As String
Private m_colValidation As Collection
Private m_bUpdated As Boolean
Private m_nPreviousRow As Integer
Private m_colKeys As Collection

Private m_ReturnCode As MSGReturnCode
' Private enums and types
Private Enum DeleteStatus
    UPDATE_DELETED
    UPDATE_NOT_DELETED
End Enum

Public m_sGroupName As String
Public m_sValueId As String
Public m_sValueName As String
Public m_sValidationType As String
Public blnSuccess As Boolean
Public m_sfrmEditComboMode As String
Public m_sfrmEditComboOperationMode As String





Public Sub SetIsEdit(Optional bIsEdit As Boolean = True)
    m_bIsEdit = bIsEdit
End Sub
Public Function IsEdit() As Boolean
    IsEdit = m_bIsEdit
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   cmdAnother_Click
' Description   :   Called when the user presses the Another button
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdAnother_Click()
    On Error GoTo Failed
    Dim bRet As Boolean
    bRet = DoOKProcessing()
    
    If bRet = True Then
        'If the record was saved, commit the transaction and begin another.
        g_clsDataAccess.CommitTrans
        g_clsDataAccess.BeginTrans
        
        g_clsFormProcessing.ClearScreenFields Me
        Form_Load
        Me.txtComboBox(GROUP_NAME_VAL).SetFocus
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   cmdCancel_Click
' Description   :   Called when the user clicks the cancel button
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdCancel_Click()
    PerformCancelProcessing
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   PerformCancelProcessing
' Description   :   Called when the user clicks the cancel button -
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PerformCancelProcessing()
    Dim clsTableAccess As TableAccess
    Dim rs As ADODB.Recordset
    
    On Error GoTo Failed
    
    Set clsTableAccess = m_clsComboValues
    Set rs = clsTableAccess.GetRecordSet
    
    ' Cancel all changes
    rs.CancelBatch
    rs.Requery
    
    CancelValidation
    
    Unload Me
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
Public Sub SetKeys(colKeys As Collection)
    Set m_colKeys = colKeys
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   cmdOK_Click
' Description   :   Called when the user pressed the OK button. Performs necessary
'                   validation and saves any data that needs to be saved
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOK_Click()
    On Error GoTo Failed
    Dim bRet As Boolean
        
    bRet = DoOKProcessing()
    
    If bRet = True Then
        SaveChangeRequest
        SetReturnCode
        Hide
    End If
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
Private Sub SaveChangeRequest()
    On Error GoTo Failed
    Dim colValues As Collection
    Dim sThisVal As Variant
    Dim sDesc As String
    Dim clsTableAccess As TableAccess
    
    Set clsTableAccess = m_clsComboGroup
    Set colValues = clsTableAccess.GetKeyMatchValues()
    
    For Each sThisVal In colValues
        sDesc = sDesc + " " + CStr(sThisVal)
    Next
    
    'BS BM0498 01/05/03
    'Remove leading blank from sDesc
    sDesc = LTrim(sDesc)
    
    g_clsHandleUpdates.SaveChangeRequest clsTableAccess, sDesc
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   CheckForDuplicates
' Description   :   Called when the user presses the OK button. Performs necessary
'                   validation to ensure no combo group or validation duplicate
'                   records exist
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CheckForDuplicates()
    
    On Error GoTo Failed
    Dim nCount As Integer
    Dim nThisRecord As Integer
    Dim clsTableAccess As TableAccess
    Dim bDuplicated As Boolean
    
    bDuplicated = g_clsFormProcessing.CheckForDuplicates(m_clsComboValues)
    
    If Not bDuplicated Then
        nCount = m_colValidation.Count
        nThisRecord = 1
        
        While nThisRecord <= nCount And bDuplicated = False
            Set clsTableAccess = m_colValidation(nThisRecord)
        
            bDuplicated = g_clsFormProcessing.CheckForDuplicates(clsTableAccess)
    
            If bDuplicated Then
                g_clsErrorHandling.RaiseError errGeneralError, "Duplicate Combo Validation records exist - please enter a unique combination"
            End If
            
            nThisRecord = nThisRecord + 1
        Wend
    Else
        g_clsErrorHandling.RaiseError errGeneralError, "Duplicate Combo Value records exist - please enter a unique combination"
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   DoOKProcessing
' Description   :   Common function used when the user presses ok or presses Another. Need
'                   to validate both the Value datagrid and the Validation datagrid. If all is
'                   ok, then perform the relevant updates to both grids.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function DoOKProcessing() As Boolean
    On Error GoTo Failed
    Dim bRet As Boolean
    Dim bFailed As Boolean
    Dim bShowError As Boolean
    Dim clsTableAccess As TableAccess
    
    bFailed = False
    bShowError = True
    
    bRet = g_clsFormProcessing.DoMandatoryProcessing(Me, bShowError)
    
    If bRet Then
        CheckForDuplicates
        
        If m_bUpdated = True And m_bIsEdit = True Then
            If MsgBox("Changes have been made to this Combo Group. If you apply these changes, there could be serious operational consequences for Omiga 4. Apply changes?", vbYesNo + vbQuestion) <> vbYes Then
                bRet = False
            End If
        End If
        
        If bRet = True Then
            bRet = dgValues.ValidateRows()
            If bRet = True Then
                bRet = dgValidation.ValidateRows()
            End If
            
            If (bRet) Then
                
                DoGroupUpdate
                
                Set clsTableAccess = m_clsComboValues
                
                'UpdateValidation UPDATE_DELETED
                
                'clsTableAccess.ValidateData
                'clsTableAccess.Update
                    
                'UpdateValidation UPDATE_NOT_DELETED
                'clsTableAccess.Update
                
                'DoUpdateComboValueAuditTableData
               
                
            End If
        Else
            bRet = False
        End If
    End If
    
    DoOKProcessing = bRet
    Exit Function

Failed:
    g_clsErrorHandling.DisplayError
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   PerformCancelProcessing
' Description   :   Called when the user clicks the cancel button - cancels
'                   all changes made to the validation table. This is because
'                   there could have been many different changes made to the
'                   validation records associated with each row of the value
'                   table.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CancelValidation() As Boolean
    Dim bRet As Boolean
    Dim nCount As Integer
    Dim nThisRecord As Integer
    Dim clsTableAccess As TableAccess
    Dim delStatus As DeleteStatus
    Dim rs As ADODB.Recordset
    On Error GoTo Failed
    
    bRet = True
    nCount = m_colValidation.Count
    nThisRecord = 1
    
    While bRet = True And nThisRecord <= nCount
        Set clsTableAccess = m_colValidation(nThisRecord)
    
        If Not clsTableAccess Is Nothing Then
            Set rs = clsTableAccess.GetRecordSet
            rs.CancelBatch
        End If
        nThisRecord = nThisRecord + 1
    Wend

    CancelValidation = bRet
    Exit Function
    
Failed:
    MsgBox "CancelValidation: Error is " + Err.DESCRIPTION
    CancelValidation = False
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   UpdateValidation
' Description   :   Updates the Validate table. For each row of the combo value table,
'                   there can be one or more validation rows. Each row may have been added,
'                   edited, or removed. If a combo value row is removed, we need to also
'                   remove the corresponding Validation rows. So, to reflect these changes,
'                   we need to update that row for it to be truly deleted. This method
'                   is called first to update all deleted rows (delStatus = UPDATE_DELETED),
'                   the again to update all newly created and edited rows (delStatus = UPDATE_NOT_DELETED).
'                   This has to be done in this order to not cause key violations.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub UpdateValidation(delStatus As DeleteStatus)
    On Error GoTo Failed
    Dim nCount As Integer
    Dim nThisRecord As Integer
    Dim clsTableAccess As TableAccess
    
    nCount = m_colValidation.Count
    nThisRecord = 1
    
    While nThisRecord <= nCount
        Set clsTableAccess = m_colValidation(nThisRecord)
    
        clsTableAccess.ValidateData
        
        If (delStatus = UPDATE_DELETED And clsTableAccess.GetIsDeleted()) Or delStatus = UPDATE_NOT_DELETED Then
            clsTableAccess.Update
        End If

        nThisRecord = nThisRecord + 1
    Wend
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
' Not used at present
Private Function ValidateScreenData() As Boolean
    Dim nCount As Integer
    Dim bRet As Boolean
    
    nCount = 0
    bRet = True
    
    ValidateScreenData = bRet
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   GetValidationKey
' Description   :   Returns the key for the validation table based on the
'                   value table key.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetValidationKey() As String
    Dim sGroup As String
    Dim sValueId As String
    Dim clsTableAccess As TableAccess
    
    On Error Resume Next
    Set clsTableAccess = m_clsComboValues
    
    sGroup = m_clsComboValues.GetGroupName()
    sValueId = m_clsComboValues.GetValueID()

    If Len(sGroup) > 0 And Len(sValueId) > 0 Then
        GetValidationKey = sGroup + sValueId
    End If
End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   GetValidationClass
' Description   :   Returns the table class for the validation table for the
'                   current value table class
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetValidationClass(clsValidation As TableAccess) As Boolean
    Dim bRet As Boolean
    Dim sKey As String
    Dim sGroup As String
    Dim sValueId As String
    Dim rs As ADODB.Recordset
    Dim clsTableAccess As TableAccess
    
    On Error Resume Next

    bRet = True
    
    sKey = GetValidationKey()
        
    If Len(sKey) > 0 Then
        Set clsValidation = m_colValidation(sKey)
        
        If Err.Number = 0 Then
            ' Already got the class loaded
            Set clsTableAccess = clsValidation
            
        ElseIf Err.Number <> 0 Then
            ' Class not loaded, so get and get it.
            Dim colMatchValues As New Collection
            Err.Clear
            
            sGroup = m_clsComboValues.GetGroupName()
            sValueId = m_clsComboValues.GetValueID()
            
            Set clsValidation = New ComboValidationTable
            m_colValidation.Add clsValidation, sKey
            
            colMatchValues.Add sGroup
            colMatchValues.Add sValueId
                        
            Set clsTableAccess = clsValidation
            clsTableAccess.SetKeyMatchValues colMatchValues
            Set rs = clsTableAccess.GetTableData()
            
        End If
    Else
        bRet = False
    End If

    GetValidationClass = bRet
    Exit Function
Failed:
    GetValidationClass = False
    MsgBox "GetValidationClass: Error is - " + Err.DESCRIPTION
End Function
'Private Sub DisplayCount()
'    Dim clsTableAccess As TableAccess
'    Dim sKey As String
'
'    On Error Resume Next
'
'    sKey = GetValidationKey()
'
'    If Len(sKey) > 0 Then
'        Set clsTableAccess = m_colValidation(sKey)
'
'        If Err.Number = 0 Then
'            Err.Clear
'            MsgBox clsTableAccess.RecordCount
'        End If
'    End If
'End Sub
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   DeleteValidation
' Description   :   Deletes all rows in the Validation table associated with
'                   key defined in the Value table. Returns the validation table
'                   class for the current row in the Value table.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub DeleteValidation()
    On Error GoTo Failed
    Dim bRet As Boolean
    Dim clsValidation As TableAccess
    Dim rs As ADODB.Recordset
    Dim colMatchValues As New Collection
    
    bRet = GetValidationClass(clsValidation)
    
    If bRet = True Then
        
        clsValidation.ValidateData
                
        If clsValidation.RecordCount() > 0 Then
            clsValidation.DeleteAllRows
        End If

    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub







Private Sub cmdValidationAdd_Click()
    
    m_sGroupName = frmEditCombo.txtComboBox(GROUP_NAME_VAL).Text
    
    If frmEditCombo.dgValidation.Rows > 0 Then
        m_sValueId = frmEditCombo.dgValidation.GetAtRowCol(dgValidation.GetCurrentRow, 1)
    Else
        m_sValueId = frmEditCombo.dgValues.GetAtRowCol(dgValues.GetCurrentRow, 1)
    End If
    m_sValidationType = ""
    m_sfrmEditComboMode = "COMBOVALIDATION"
    m_sfrmEditComboOperationMode = "ADD"
    
    frmNewCombo.Show vbModal, Me
    
    If blnSuccess = True Then
        Form_Load
    End If
    
End Sub

Private Sub cmdValidationEdit_Click()
        
    m_sGroupName = frmEditCombo.txtComboBox(GROUP_NAME_VAL).Text
    
    If frmEditCombo.dgValidation.Rows > 0 Then
        m_sValueId = frmEditCombo.dgValidation.GetAtRowCol(dgValidation.GetCurrentRow, 1)
    Else
        m_sValueId = frmEditCombo.dgValues.GetAtRowCol(dgValues.GetCurrentRow, 1)
    End If
    
    m_sValidationType = frmEditCombo.dgValidation.GetAtRowCol(dgValidation.GetCurrentRow, 2)
    
    m_sfrmEditComboMode = "COMBOVALIDATION"
    m_sfrmEditComboOperationMode = "EDIT"
    
    frmNewCombo.Show vbModal, Me
    
    If blnSuccess = True Then
        Form_Load
        
    End If
End Sub

Private Sub cmdValueAdd_Click()
    
    m_sGroupName = frmEditCombo.txtComboBox(GROUP_NAME_VAL).Text
    m_sValueId = ""
    m_sValueName = ""
    m_sfrmEditComboMode = "COMBOVALUE"
    m_sfrmEditComboOperationMode = "ADD"
    frmNewCombo.Show vbModal, Me
    
    If blnSuccess Then
        Form_Load
        m_bIsEdit = True
    End If

End Sub

Private Sub cmdValueEdit_Click()
   
    m_sGroupName = frmEditCombo.txtComboBox(GROUP_NAME_VAL).Text
    m_sValueId = frmEditCombo.dgValues.GetAtRowCol(dgValues.GetCurrentRow, 1)
    m_sValueName = frmEditCombo.dgValues.GetAtRowCol(dgValues.GetCurrentRow, 2)
    m_sfrmEditComboMode = "COMBOVALUE"
    m_sfrmEditComboOperationMode = "EDIT"
    frmNewCombo.Show vbModal, Me
    
    If blnSuccess Then
        Form_Load
        SetupValidationGrid
    End If
    
End Sub

Private Sub dgValidation_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
    
    m_bUpdated = True
   
    If ColIndex = 2 Then
        g_clsErrorHandling.DisplayError "You cannot edit the ValidationType"
        Cancel = True
    End If
    
End Sub






Private Sub dgValues_BeforeDelete(Cancel As Integer)
    DeleteValidation
End Sub
Private Sub dgValues_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
    
    Dim field As FieldData
    
    If ColIndex = 1 Or ColIndex = 2 Then
        g_clsErrorHandling.DisplayError "You cannot edit the ValueId"
        Cancel = True
        Exit Sub
    Else
        field = dgValues.GetFieldData(ColIndex)
        If field.sTitle = VALUE_TITLE Then
            frmEditCombo.dgValues.Columns(1).Locked = True
            Cancel = DoesChildRecordExist()
    
            If Cancel <> 0 Then
                g_clsErrorHandling.DisplayError "Validation record exists for this record - remove Validation record before edit"
            Else
                m_bUpdated = True
            End If
        End If
    End If
     
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   DoesChildRecordExist
' Description   :   Determines whether or not, for the current recvord, if a child
'                   validation row exists.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function DoesChildRecordExist()
    Dim bRet As Boolean
    Dim sGroup As String
    Dim sValueId As String
    Dim rs As ADODB.Recordset
    Dim clsTableAccess As TableAccess
    Dim colMatchValues As New Collection
    
    bRet = False
    Set clsTableAccess = New ComboValidationTable
    
    sGroup = m_clsComboValues.GetGroupName()
    
    'BMIDS00565 08/10/02 CL
    If Len(sGroup) > 0 Then
        sValueId = m_clsComboValues.GetValueID()
        If Len(sGroup) > 0 And Len(sValueId) > 0 Then
            colMatchValues.Add sGroup
            colMatchValues.Add sValueId
        
            clsTableAccess.SetKeyMatchValues colMatchValues
    
            Set rs = clsTableAccess.GetTableData()
    
            If Not rs Is Nothing Then
                If rs.RecordCount > 0 Then
                    bRet = True
                End If
            End If
        End If
    'BMIDS00565 08/10/02 CL END
    End If
    DoesChildRecordExist = bRet
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   dgValues_RowColChange
' Description   :   When a row in the Value table (Datagrid) changes, we need
'                   to populate the Validation datagrid with the validation data
'                   associated with that row.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub dgValues_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
        
    SetupValidationGrid
    
    If dgValidation.Rows > 0 Then
        frmEditCombo.cmdValidationEdit.Enabled = True
    Else
        frmEditCombo.cmdValidationEdit.Enabled = False
    End If
    
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   Form_Load
' Description   :   Initialisation to this screen
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()

    On Error GoTo Failed
    Dim colMatchValues As New Collection
    Dim clsTableAccess As TableAccess
    
    SetReturnCode MSGFailure
    m_bUpdated = False
    
    Set m_colValidation = New Collection
    Set m_clsComboValues = New ComboValueTable
    Set m_clsComboValidation = New ComboValidationTable
    Set m_clsComboGroup = New ComboValueGroupTable
    
    If m_bIsEdit = True Then
        SetEditState
    Else
        SetAddState
    End If

    m_sGroupNameOrig = txtComboBox(GROUP_NAME_VAL).Text
    m_sGroupNoteOrig = txtComboNotes.Text
    
    If dgValues.Rows > 0 Then
        cmdValueEdit.Enabled = True
        cmdValidationAdd.Enabled = True
    Else
        cmdValueEdit.Enabled = False
        cmdValidationAdd.Enabled = False
        cmdValidationEdit.Enabled = False
    End If
    
    If dgValidation.Rows > 0 Then
        cmdValidationEdit.Enabled = True
    Else
        cmdValidationEdit.Enabled = False
    End If
        
    m_sfrmEditComboMode = ""
    m_sGroupName = ""
    m_sValueId = ""
    m_sValueName = ""
    m_sValidationType = ""
    blnSuccess = False
   
    Exit Sub

Failed:
    g_clsErrorHandling.DisplayError
    g_clsFormProcessing.CancelForm Me
End Sub
Public Sub SetAddState()
    On Error GoTo Failed
    SetupDBControl m_clsComboValues, dgValues
    SetupValidationGrid
    cmdAnother.Enabled = True
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Public Sub SetEditState()
    Dim clsTableAccess As TableAccess
    
    dgValues.Enabled = True
        
    Set clsTableAccess = m_clsComboValues
    clsTableAccess.SetKeyMatchValues m_colKeys
    
    Set clsTableAccess = m_clsComboGroup
    clsTableAccess.SetKeyMatchValues m_colKeys
    
    SetupDBControl m_clsComboValues, dgValues
    
    If dgValues.Rows > 0 Then
        dgValues.Row = 0
        dgValues.SelBookmarks.Add adBookmarkFirst
        SetupValidationGrid
    End If
    Me.txtComboBox(GROUP_NAME_VAL).Enabled = False
    cmdAnother.Enabled = False
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set m_colValidation = Nothing
    ' m_blnAddNewComboRecord = False
    ' m_blnNewValidationRecord = False
End Sub

Private Sub txtComboBox_Validate(Index As Integer, Cancel As Boolean)
    On Error GoTo Failed
    Cancel = Not Me.txtComboBox(Index).ValidateData()
    
    If Cancel = False Then
        DoesGroupExist
        SetValueDataGridState
    End If
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
    Cancel = True
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SetupDBControl
' Description   :   Sets up the datagrid for the table and grid passed in.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetupDBControl(clsTableAccess As TableAccess, dbControl As Object)
    On Error GoTo Failed
    Dim colTables As New Collection
    Dim colDataControls As New Collection
    
    colTables.Add clsTableAccess
    colDataControls.Add dbControl
    
    If m_bIsEdit = True Then
        g_clsFormProcessing.PopulateDBScreen colTables, colDataControls
    Else
        g_clsFormProcessing.PopulateDBScreen colTables, colDataControls, POPULATE_EMPTY
    End If
    
    ' If we're editing we need to read the text edit values too.
    If m_bIsEdit = True Then
        PopulateScreenFields
    End If

    SetValueGridFields

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SetValueGridFields
' Description   :   Sets the field names for the Value grid, specifies which
'                   are mandatory and which are banded.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetValueGridFields()
    Dim fields As FieldData
    Dim colFields As New Collection
    Dim clsTableAccess As TableAccess
    Dim sField As String
    
    ' First, Group Name. Not visible
    fields.sField = "GroupName"
    fields.bRequired = True
    fields.bVisible = False
    fields.sDefault = txtComboBox(GROUP_NAME_VAL).Text
    fields.sError = ""
    fields.sTitle = ""
    colFields.Add fields
    
    ' StartDate not visible, but has to be copied in.
    fields.sField = "ValueID"
    fields.bRequired = True
    fields.bVisible = True
    fields.sDefault = ""
    fields.sError = "Value ID must be entered"
    fields.sTitle = VALUE_TITLE
    colFields.Add fields

    ' Next, TypeOfApplication has to be entered
    fields.sField = "ValueName"
    fields.bRequired = True
    fields.bVisible = True
    fields.sDefault = ""
    fields.sError = "Value Name must be entered"
    fields.sTitle = "Value Name"
    colFields.Add fields
    
    Set clsTableAccess = m_clsComboValues
    
    sField = "OM3VALUEID"
    
    If g_clsDataAccess.DoesFieldExist(clsTableAccess.GetTable(), sField, , False) Then
        ' Next, Omiga 3 Value ID
        fields.sField = sField
        fields.bRequired = False
        fields.bVisible = False
        fields.sDefault = ""
        fields.sError = ""
        fields.sTitle = ""
        colFields.Add fields
    End If
    
    dgValues.SetColumns colFields, "EditComboValue", "Combo Values"

End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SetValidationGridFields
' Description   :   Sets the field names for the Validation grid, specifies which
'                   are mandatory and which are banded.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetValidationGridFields(sGroup As String, sValueId As String)
    Dim fields As FieldData
    Dim colFields As New Collection
    
    ' First, Group Name. Not visible
    fields.sField = "GroupName"
    fields.bRequired = True
    fields.bVisible = False
    fields.sDefault = sGroup
    fields.sError = ""
    fields.sTitle = ""
    colFields.Add fields
    
    ' StartDate not visible, but has to be copied in.
    fields.sField = "ValueID"
    fields.bRequired = True
    fields.bVisible = False
    fields.sDefault = sValueId
    fields.sError = ""
    fields.sTitle = ""
    colFields.Add fields
    
    ' Next, TypeOfApplication has to be entered
    fields.sField = "ValidationType"
    fields.bRequired = True
    fields.bVisible = True
    fields.sDefault = ""
    fields.sError = "Validation Type must be entered"
    fields.sTitle = "Validation Type"
    colFields.Add fields
    
    dgValidation.SetColumns colFields, "EditComboValidation", "Combo Validation"
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   PopulateGroupFields
' Description   :   For the current combo Group, read the group name and start
'                   dates and set the corresponding text fields
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PopulateGroupFields()
    On Error GoTo Failed
    Dim colMatchValues As Collection
    Dim clsTableAccess As TableAccess
    Dim rs As ADODB.Recordset
    
    On Error GoTo Failed

    Set clsTableAccess = m_clsComboGroup
    
    Set rs = clsTableAccess.GetTableData()
    clsTableAccess.ValidateData
        
    If rs.RecordCount > 0 Then
        txtComboBox(GROUP_NAME_VAL).Text = m_clsComboGroup.GetGroupName()
        txtComboNotes.Text = m_clsComboGroup.GetGroupNote()
    Else
        g_clsErrorHandling.RaiseError errRecordNotFound, "ComboGroup"
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Public Sub PopulateScreenFields()
    On Error GoTo Failed
    Dim bRet As Boolean
    Dim clsComboValues  As ComboValueTable
    Dim colKeys As New Collection
        
    PopulateGroupFields
        
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'Private Sub DoComboUpdates()
'    Dim bRet As Boolean
'    Dim sGroup As String
'    Dim sValueID As String
'    Dim clsBandedTable As BandedTable
'    Dim clsTableAccess As TableAccess
'
'    On Error GoTo Failed
'    ' If the Name, Description, or Date has changed then we need to update ALL rows
'    ' with the new Values.
'    sGroup = Me.txtComboBox(GROUP_NAME_VAL).Text
'
'    bRet = True
'
'    If m_sGroupNameOrig <> sGroup Then
'        Dim colValues As New Collection
'
'        ' First the Validation Table
'        bRet = GetValidationClass(clsTableAccess)
'
'        If bRet = True Then
'            sValueID = m_clsComboValues.GetValueID()
'
'            If Len(sValueID) > 0 Then
'                colValues.Add sGroup
'                colValues.Add sValueID
'                Set clsBandedTable = clsTableAccess
'                clsBandedTable.SetUpdateValues colValues
'
'                ' We only want to check that the record exists if the name or start date has changed
'                clsBandedTable.SetUpdateSets
'
'                clsBandedTable.DoUpdateSets
'
'                ' Need to update the m_colValidation collection
'                Dim sKey As String
'
'                sKey = m_sGroupNameOrig + sValueID
'
'                If Len(sKey) > 0 Then
'                    m_colValidation.Remove sKey
'
'                    sKey = sGroup + sValueID
'
'                    MsgBox clsTableAccess.RecordCount
'                    m_colValidation.Add clsTableAccess, sKey
'                Else
'                    g_clsErrorHandling.RaiseError errGeneralError, "EditCombo:DoComboUpdates - Key is empty"
'                End If
'            Else
'                g_clsErrorHandling.RaiseError errGeneralError, "Cannot update validation - no combo value selected"
'            End If
'        End If
'
'        ' Point at a BandedGlobalPArameters object to set the update values
'        Set clsBandedTable = m_clsComboValues
'        Set colValues = New Collection
'        colValues.Add sGroup
'
'        clsBandedTable.SetUpdateValues colValues
'
'        ' We only want to check that the record exists if the name or start date has changed
'        clsBandedTable.SetUpdateSets
'
'        clsBandedTable.DoUpdateSets
'        DisplayCount
'        UpdateValueCombo
'    End If
'
'    Exit Sub
'Failed:
'    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
'End Sub
'Private Sub UpdateValueCombo()
'    On Error GoTo Failed
'    Dim bRet As Boolean
'    Dim rsValueOrig As ADODB.Recordset
'    Dim rsCopyOrig As ADODB.Recordset
'    Dim clsTableAccess As TableAccess
'    Dim clsCopy As New ComboValueTable
'    Dim clsValidation As ComboValidationTable
'    Dim clsValidationCopy As New ComboValidationTable
'
'    Set clsTableAccess = m_clsComboValues
'    Set rsValueOrig = clsTableAccess.GetRecordSet()
'    DisplayCount
'
'    clsTableAccess.ValidateData
'
'    If rsValueOrig.RecordCount > 0 Then
'        g_clsFormProcessing.CopyRecordset clsTableAccess, clsCopy
'
'        DisplayCount
'        ' Copy performed, so delete the original
'        clsTableAccess.DeleteAllRows
'        ' Now use the copy
'        Set clsTableAccess = clsCopy
'        Set rsCopyOrig = clsTableAccess.GetRecordSet()
'        Set clsTableAccess = m_clsComboValues
'        clsTableAccess.SetRecordSet rsCopyOrig
'        Set dgValues.DataSource = rsCopyOrig
'
'        ' Now need to copy the validation table and delete the original
'        bRet = GetValidationClass(clsValidation)
'
'        If bRet = True Then
'            Set clsTableAccess = clsValidation
'
'            g_clsFormProcessing.CopyRecordset clsValidation, clsValidationCopy
'
'            clsTableAccess.DeleteAllRows
'
'            Set clsTableAccess = clsValidationCopy
'            Set rsCopyOrig = clsTableAccess.GetRecordSet()
'            Set clsTableAccess = clsValidation
'            clsTableAccess.SetRecordSet rsCopyOrig
'
'            ' Now Update the Validation record
'            clsTableAccess.Update
'
'            Set clsTableAccess = m_clsComboValues
'            clsTableAccess.Update
'        End If
'    End If
'
'    Exit Sub
'Failed:
'    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
'End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   DoesGroupExist
' Description   :   Determines if the group comprised of the start date and
'                   group name already exists in the combo group table.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub DoesGroupExist()
    On Error GoTo Failed
        
    Dim bExists As Boolean
    Dim sGroup As String
    Dim col As New Collection
    Dim clsTableAccess As TableAccess
    
    sGroup = txtComboBox(GROUP_NAME_VAL).Text
    Set clsTableAccess = m_clsComboGroup
    
    If Len(sGroup) > 0 Then
        col.Add sGroup
        clsTableAccess.SetKeyMatchValues col
        
        bExists = Not clsTableAccess.DoesRecordExist(col)
                
        If bExists = False Then
            Me.txtComboBox(GROUP_NAME_VAL).SetFocus
            g_clsErrorHandling.RaiseError errGeneralError, "Entry for " + sGroup + " already exists"
        End If
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   DoGroupUpdate
' Description   :   Updates the group table with the group name and start
'                   date the user has entered. Only does this if the group
'                   doesn't already exist.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub DoGroupUpdate()
    On Error GoTo Failed
    Dim bRet As Boolean
    Dim sGroup As String
    Dim sNotes As String
    Dim rs As ADODB.Recordset
    Dim clsTableAccess As TableAccess
    Dim bUpdate As Boolean
    Dim col As New Collection
        
    On Error GoTo Failed
    sGroup = txtComboBox(GROUP_NAME_VAL).Text
    sNotes = txtComboNotes.Text
    bUpdate = False
    
    Set clsTableAccess = m_clsComboGroup
    
    If Len(sGroup) > 0 Then
        ' Do the group duplicate search if adding (it's new) or if it's changed
        If m_bIsEdit = False Or (m_bIsEdit = True And m_sGroupNameOrig <> sGroup) Then
            DoesGroupExist
        End If
                
        If m_bIsEdit = False Then
            ' Adding, so add a new record
            g_clsFormProcessing.CreateNewRecord clsTableAccess
        End If
        
        m_clsComboGroup.SetGroupName sGroup
        m_clsComboGroup.SetGroupNote sNotes
        bUpdate = True
    End If
    
    If bUpdate = True Then
        clsTableAccess.Update
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'Private Function DoGroupUpdate() As Boolean
'    Dim bRet As Boolean
'    Dim sGroup As String
'    Dim sNotes As String
'    Dim rs As ADODB.Recordset
'    Dim clsTableAccess As TableAccess
'    Dim bUpdate As Boolean
'    Dim col As New Collection
'
'    sGroup = txtComboBox(GROUP_NAME_VAL).Text
'    sNotes = txtComboNotes.Text
'    bUpdate = False
'
'    Set clsTableAccess = m_clsComboGroup
'
'    If Len(sGroup) > 0 Then
'        If m_bisedit = True Then
'            If sNotes <> m_sGroupNoteOrig Then
'                m_clsComboGroup.SetGroupNote sNotes
'                bUpdate = True
'            End If
'            bRet = True
'        Else
'            col.Add sGroup
'            clsTableAccess.SetKeyMatchValues col
'
'            bRet = clsTableAccess.DoesRecordExist(col)
'
'            If bRet = False Then
'                ' Doesn't exist, so add a new one
'                bRet = g_clsFormProcessing.CreateNewRecord(clsTableAccess)
'
'                If bRet = True Then
'                    m_clsComboGroup.SetGroupName sGroup
'                    m_clsComboGroup.SetGroupNote sNotes
'                    bUpdate = True
'                Else
'                    MsgBox "EditComboValues:DoGroupUpdate - Unable to create new record"
'                End If
'            Else
'                MsgBox "Entry for " + sGroup + " already exists"
'                Me.txtComboBox(GROUP_NAME_VAL).SetFocus
'                bRet = False
'            End If
'        End If
'    End If
'
'    If bUpdate = True Then
'        bRet = clsTableAccess.Update()
'    End If
'
'    DoGroupUpdate = bRet
'    Exit Function
'Failed:
'    MsgBox "EditAdminFees:DoGroupUpdate - Error is " + Err.DESCRIPTION
'    DoGroupUpdate = False
'End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SetupValidationGrid
' Description   :   Populates the validation grid based on the keys obtained
'                   from the combo value datagrid.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetupValidationGrid()
    On Error GoTo Failed
    Dim bPopulateEmpty As Boolean
    Dim sGroup As String
    Dim sValueId As String
    Dim sKey As String
    Dim clsTableAccess As TableAccess
    Dim colMatchValues As New Collection
    Dim rs As ADODB.Recordset
    
    Set clsTableAccess = m_clsComboValues
    On Error Resume Next
    bPopulateEmpty = True
    
    ' DJP SQL Server port - check there's a record before trying to read any values
    If clsTableAccess.RecordCount > 0 Then
        
        sGroup = m_clsComboValues.GetGroupName()
        sValueId = m_clsComboValues.GetValueID()

        If Len(sGroup) > 0 And Len(sValueId) > 0 Then
            colMatchValues.Add sGroup
            colMatchValues.Add sValueId
            
            sKey = sGroup + sValueId
            Set clsTableAccess = m_colValidation.Item(sKey)
            
            If Err.Number = 0 Then
                Set rs = clsTableAccess.GetRecordSet
            Else
                Err.Clear
                Set clsTableAccess = New ComboValidationTable
                m_colValidation.Add clsTableAccess, sKey
                clsTableAccess.SetKeyMatchValues colMatchValues
                Set rs = clsTableAccess.GetTableData()
            End If
    
            If Not rs Is Nothing Then
                Set dgValidation.DataSource = rs
                dgValidation.Enabled = True
                SetValidationGridFields sGroup, sValueId
                
            End If
            bPopulateEmpty = False
        End If
    End If
    
    If bPopulateEmpty Then
        ' Just populate the grid with nothing so we can show the fields
        Set clsTableAccess = New ComboValidationTable
        Set rs = clsTableAccess.GetTableData(POPULATE_EMPTY)
        Set dgValidation.DataSource = rs
        SetValidationGridFields "", ""
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SetValueDataGridState
' Description   :   Enables the value datagrid if the required values are
'                   entered - the mandatory data such as the start date
'                   and group name
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetValueDataGridState()
    Dim bRet As Boolean
    Dim bShowError As Boolean
    bShowError = False
    
    bRet = g_clsFormProcessing.DoMandatoryProcessing(Me, bShowError)
    
    If bRet = True Then
        Dim bEnabled As Boolean
        
        bEnabled = dgValues.Enabled
        dgValues.Enabled = True
        
        SetValueGridFields
        
        If m_bIsEdit = False Then
            If bEnabled = False Then
                dgValues.AddRow
            End If
        End If
    End If
End Sub
Private Sub SetReturnCode(Optional enumReturn As MSGReturnCode = MSGSuccess)
    m_ReturnCode = enumReturn
End Sub
Public Function GetReturnCode() As MSGReturnCode
    GetReturnCode = m_ReturnCode
End Function

Private Sub txtComboNotes_LostFocus()
    DoGroupUpdate
End Sub
