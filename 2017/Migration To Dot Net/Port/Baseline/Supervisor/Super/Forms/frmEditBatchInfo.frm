VERSION 5.00
Object = "{24A7A369-0EDB-4924-817E-D81FE8E09890}#1.0#0"; "MSGOCX.ocx"
Begin VB.Form frmBatchSchedule 
   Caption         =   "Batch Schedule Info"
   ClientHeight    =   6060
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10380
   Icon            =   "frmEditBatchInfo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6060
   ScaleWidth      =   10380
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdViewErrors 
      Caption         =   "&View Errors"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   3600
      TabIndex        =   8
      Top             =   5520
      Width           =   1455
   End
   Begin VB.CommandButton cmdRunFailedBatchJob 
      Caption         =   "&Run Failed Batch"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   1920
      TabIndex        =   7
      Top             =   5520
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   8920
      TabIndex        =   6
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel Batch"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   5520
      Width           =   1455
   End
   Begin MSGOCX.MSGListView lvBatchInstance 
      Height          =   4635
      Left            =   195
      TabIndex        =   4
      Top             =   720
      Width           =   10020
      _ExtentX        =   17674
      _ExtentY        =   8176
      Sorted          =   -1  'True
      AllowColumnReorder=   0   'False
   End
   Begin MSGOCX.MSGMulti txtBatchNumber 
      Height          =   315
      Left            =   1440
      TabIndex        =   2
      Top             =   180
      Width           =   1275
      _ExtentX        =   2249
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
      Text            =   ""
   End
   Begin MSGOCX.MSGMulti txtProgramType 
      Height          =   315
      Left            =   4440
      TabIndex        =   3
      Top             =   180
      Width           =   1995
      _ExtentX        =   3519
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
      Text            =   ""
   End
   Begin VB.Label lblProgramType 
      AutoSize        =   -1  'True
      Caption         =   "Program Type"
      Height          =   195
      Left            =   3240
      TabIndex        =   1
      Top             =   240
      Width           =   990
   End
   Begin VB.Label lblBatchNumber 
      AutoSize        =   -1  'True
      Caption         =   "Batch Number"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1020
   End
End
Attribute VB_Name = "frmBatchSchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form          : frmEditBatchInfo
' Description   : Form which allows the adding/editing of batch programs
'
' Change history
' Prog      Date        Description
' AA        13/02/01    Added Form
' SA        14/02/02    SYS4073 New button Added to view batch errors
'                       Display new error fields from Batch Schedule record
' GHun      08/04/2002  SYS4368 Errors in failed batch run, re-run
' GHun      17/04/2002  SYS4410 Supervisor may crash when running a failed batch
' GHun      18/04/02    SYS4424 In Supervisor you cant view errors in incomplete batch runs
' CL        17/05/2002  SYS4557 Revision to viewing errors
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' BMids change history
' Prog      Date        Description
' GHun      28/03/2003  BM0425 Prevent unnecessary locks from being created & aesthetic improvements
' JD        30/11/2004  BMIDS604    on OK from the screen, refresh the view of the batch jobs in order to see the latest status
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Epsom change history
' Prog      Date        Description
' TW        19/02/2007  EP2_1348 - Error in certain circumstances on MSGListView when AllowColumnReorder set to true
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
'Private Const PROGRAMTYPE_COMBO As String = "BatchProgramType"
Private Const m_sStatusComboGroup As String = "BatchScheduleStatus"
Private Const m_sTableName As String = "BATCHSCHEDULE"
Private Const m_nBatchRunNoColumn = 1
Private Const m_nStatusColumn = 2

'SYS4073 Set up constants for all batch schedule status
Private Const m_sStatusCWEValidType = "CWE"
Private Const m_sStatusComplete = "C"
Private Const m_sStatusRunning = "R"
Private Const m_sStatusWaiting = "W"
Private Const m_sStatusCancelled = "CA"

Private m_ReturnCode As MSGReturnCode
Private m_clsBatchSchedule As BatchScheduleTable
'Private m_clsBatch As BatchTable
Private m_colKeys  As Collection
Private m_sBatchNumber As String
Private m_bIsEdit As Boolean
Private m_clsTableAccess As TableAccess
Private m_sBatchType As String          'BM0425

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SetIsEdit
' Description   :   Sets whether or not this form is in edit or add mode. Defaults to edit
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetIsEdit(Optional bIsEdit As Boolean = True)
    m_bIsEdit = bIsEdit
End Sub

Private Sub SetReturnCode(Optional enumReturn As MSGReturnCode = MSGSuccess)
    m_ReturnCode = enumReturn
End Sub

Public Function GetReturnCode() As MSGReturnCode
    GetReturnCode = m_ReturnCode
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SetListViewHeaders
' Description   :   Sets the field header titles on the listview
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetListViewHeaders()
    On Error GoTo Failed
    Dim headers As New Collection
    Dim lvHeaders As listViewAccess
    
    lvHeaders.nWidth = 15
    lvHeaders.sName = "Batch Run No"
    headers.Add lvHeaders
    
    lvHeaders.nWidth = 20
    lvHeaders.sName = "Status"
    headers.Add lvHeaders
    
    lvHeaders.nWidth = 10
    lvHeaders.sName = "Created On"
    headers.Add lvHeaders
    
    lvHeaders.nWidth = 9
    lvHeaders.sName = "Completed On"
    headers.Add lvHeaders
    
    lvHeaders.nWidth = 9
    lvHeaders.sName = "Executed On"
    headers.Add lvHeaders
    
    lvHeaders.nWidth = 10
    lvHeaders.sName = "No Records"
    headers.Add lvHeaders
    
    lvHeaders.nWidth = 10
    lvHeaders.sName = "No Successes"
    headers.Add lvHeaders
    
    lvHeaders.nWidth = 10
    lvHeaders.sName = "No Failures"
    headers.Add lvHeaders
    
    'SYS4073 ErrorNo/ErrorSource/ErrorDescription
    lvHeaders.nWidth = 10
    lvHeaders.sName = "Error No"
    headers.Add lvHeaders
    
    lvHeaders.nWidth = 25
    lvHeaders.sName = "Error Source"
    headers.Add lvHeaders
        
    lvHeaders.nWidth = 35
    lvHeaders.sName = "Error Description"
    headers.Add lvHeaders
    
    lvBatchInstance.AddHeadings headers
    lvBatchInstance.LoadColumnDetails TypeName(Me)
    
    lvBatchInstance.HideSelection = False
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cmdCancel_Click
' Description   : Cancels the selected batch item
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdCancel_Click()
        
    On Error GoTo Failed
    Dim sBatchrunNo As String
    Dim bRet As Boolean
    Dim nSelectedItem As Long
    lvBatchInstance.GetSelectedColumn m_nBatchRunNoColumn, sBatchrunNo
    
    bRet = CancelSelectedBatch(sBatchrunNo)
    
    nSelectedItem = lvBatchInstance.SelectedItem.Index
    
    'BM0425 PopulateListView should always be called
    'If bRet Then
    PopulateListView
    'End If
    'BM0425 End
        
    lvBatchInstance.ListItems.Item(nSelectedItem).Selected = True
    lvBatchInstance.SetFocus
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError Err.DESCRIPTION
    lvBatchInstance.SetFocus
End Sub

Private Sub cmdOK_Click()
    On Error GoTo Failed
    g_clsMainSupport.GetBatch frmMain.lvListView 'JD BMIDS604 refresh the list
    Hide
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError Err.DESCRIPTION
End Sub


Private Sub cmdRunFailedBatchJob_Click()
    On Error GoTo Failed
    
    RunFailedBatch
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError Err.DESCRIPTION
End Sub

Private Sub cmdViewErrors_Click()
    Dim frmErrors As frmBatchErrors
    Dim sBatchrunNo As String
    Dim sBatchStatus As String
    Dim colBatchAuditKeys As Collection
    Dim sSQL As String
       
    Set colBatchAuditKeys = New Collection
    'Add Batch Number
    colBatchAuditKeys.Add m_colKeys(1)
    
    ' Get the BatchRunNo
    lvBatchInstance.GetSelectedColumn m_nBatchRunNoColumn, sBatchrunNo
    ' Add the Batch Run No
    colBatchAuditKeys.Add sBatchrunNo
    
    Set frmErrors = New frmBatchErrors
    
    frmErrors.SetKeys colBatchAuditKeys
    'BM0425
    'frmErrors.SetBatchType m_clsBatchSchedule.GetProgValidationType
    frmErrors.SetBatchType m_sBatchType
    'BM0425 End
                    
    frmErrors.Show vbModal

    Unload frmErrors
    Set frmErrors = Nothing
    
    Exit Sub
End Sub

'SYS4424 Pressing F5 should refresh the list view.
'This requires the KeyPreview property on the form to be set to true
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo Failed
    
    If KeyCode = vbKeyF5 Then
        BeginWaitCursor
        m_clsBatchSchedule.GetBatchScheduleTableData m_sBatchNumber
        PopulateListView
        KeyCode = 0
        EndWaitCursor
    End If
    Exit Sub
Failed:
    EndWaitCursor
    g_clsErrorHandling.DisplayError Err.DESCRIPTION
End Sub
'SYS4424 End

Private Sub Form_Load()
    On Error GoTo Failed
    
    Set m_clsBatchSchedule = New BatchScheduleTable
    Set m_clsTableAccess = m_clsBatchSchedule
    
    m_sBatchNumber = CStr(m_colKeys(1))
    
    PopulateScreen
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError Err.DESCRIPTION
    g_clsFormProcessing.CancelForm Me
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : PopulateScreen
' Description   : Check if there are any Schedule records, and if there are then populates the screen
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PopulateScreen()

    On Error GoTo Failed
    'BM0425 Not used
    'Dim clsTableAccess As TableAccess
    'Dim rs As ADODB.Recordset
    'Dim vTmp As Variant
    
    'Set clsTableAccess = m_clsBatchSchedule
    'clsTableAccess.SetKeyMatchValues m_colKeys
    'BM0425 End
    
    'BM0425 returned recordset is never used
    'Set rs = m_clsBatchSchedule.GetBatchScheduleTableData(m_sBatchNumber)
    
    'BM0425 GetBatchScheduleTableData is already called when populating the list view
    '       so this call should be avoided to prevent querying the same data twice
    'm_clsBatchSchedule.GetBatchScheduleTableData m_sBatchNumber
    
    'BM0425 Moved the RecordCount Check and Error message to PopulateListView
    'Are the any Schedule records for this batch process?
    'If clsTableAccess.RecordCount > 0 Then
        PopulateScreenControls
        PopulateScreenFields
    'Else
    '    g_clsErrorHandling.RaiseError errGeneralError, "There are no Schedules Batch Jobs for " & m_sBatchNumber
    'End If
    'BM0425 End

    Exit Sub

Failed:
     g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Public Sub SetKeys(colKeys As Collection)
    Set m_colKeys = colKeys
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : PopulateScreenFields
' Description   : Populate all fields on the form
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PopulateScreenFields()

    On Error GoTo Failed
    
    PopulateListView
    'SYS4073
    'BM0425 SetButtonState moved into PopulateListView
    'SetButtonState
    txtBatchNumber.Text = m_clsBatchSchedule.GetBatchNumber
    txtProgramType.Text = m_clsBatchSchedule.GetProgramTypeText
    m_sBatchType = m_clsBatchSchedule.GetProgValidationType
    'BM0425 End
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : CancelSelectedBatch
' Description   : Cancels the selected batch
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CancelSelectedBatch(sBatchrunNo As String) As Boolean
    On Error GoTo Failed
    
    Dim bRet As Boolean
    'Dim clsTableAccess As TableAccess
    'Dim sSQL As String
    
    'BM0425 Query the specific row to be updated
    m_clsBatchSchedule.GetBatchScheduleForUpdate m_sBatchNumber, sBatchrunNo
    
    'sSQL = "BatchRunNumber = " & sBatchRunNo
    '    Set clsTableAccess = m_clsBatchSchedule
    '
    ''Apply filter so the current record is the one selected
    'clsTableAccess.ApplyFilter sSQL
    'BM0425 End
    bRet = m_clsBatchSchedule.CancelBatchRecord
    
    'clsTableAccess.CancelFilter
 
    CancelSelectedBatch = bRet
    
    Exit Function
Failed:
    'BM0425
    CancelSelectedBatch = False
    g_clsErrorHandling.DisplayError Err.DESCRIPTION
    lvBatchInstance.SetFocus
    'g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
    'BM0425 End
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : PopulateScreenControls
' Description   : Sets the listview Headers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PopulateScreenControls()
    On Error GoTo Failed
    
    SetListViewHeaders
    txtBatchNumber.Enabled = False
    txtProgramType.Enabled = False
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : PopulateListView
' Description   : Populates the listview with all schedule records
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PopulateListView()
    
    On Error GoTo Failed
    
    'BM0425
    'g_clsMainSupport.GetBatchSchedule lvBatchInstance, m_sBatchNumber
    m_clsBatchSchedule.GetBatchScheduleTableData m_sBatchNumber

    If m_clsTableAccess.RecordCount > 0 Then
        g_clsFormProcessing.PopulateFromRecordset lvBatchInstance, m_clsBatchSchedule
        SetButtonState
    Else
        g_clsErrorHandling.RaiseError errGeneralError, "There are no Scheduled Batch Jobs for " & m_sBatchNumber
    End If
    'BM0425 End
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : RunFailedBatch
' Description   : Sets the status to waiting and adds the schedule recors to the queue
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RunFailedBatch()
    On Error GoTo Failed
    
    Dim bRet As Boolean
    'Dim clsTableAccess As TableAccess
    'Dim sSQL As String
    Dim sBatchrunNo As String
    Dim clsBatch As BatchProcessing
    Dim dExecution As Date
    Dim nSelectedItem As Long
    'Dim nStatusRunning As Integer
    'Dim clsComboValidation As ComboValidationTable
    
    Set clsBatch = New BatchProcessing
    'Set clsComboValidation = New ComboValidationTable
    
    lvBatchInstance.GetSelectedColumn m_nBatchRunNoColumn, sBatchrunNo
    
    If lvBatchInstance.SelectedItem.Index <= 0 Then
        g_clsErrorHandling.RaiseError errGeneralError, "You must select a Batch Schedule"
    End If
    
    'BM0425 Get the data so that it can be updated
    BeginWaitCursor
    m_clsBatchSchedule.GetBatchScheduleForUpdate m_sBatchNumber, sBatchrunNo
    'sSQL = "BatchRunNumber = " & sBatchRunNo
    'Set clsTableAccess = m_clsBatchSchedule
    
    'Apply filter so the current record is the one selected
    'clsTableAccess.ApplyFilter sSQL
    'BM0425 End

    If m_clsBatchSchedule.GetBatchRetryIndicator = 0 Then
        'Get the execution date
        dExecution = m_clsBatchSchedule.GetBatchExecutionDate
        'Get waiting Status id from combo group
        
        'SYS4410 nStatusRunning is no longer used so it has been commented out
        'nStatusRunning = clsComboValidation.GetValueIDForValidationType(BATCH_SCHEDULE_STATUS_WAITING, m_sStatusComboGroup)
        
        'Launch the batch schedule record
        ' AA removed.
        'SYS3327 This needs to go back in to call the m/tier.
        'SYS4073 Send in indicator to tell its a re-run.
        'SYS4368 Pass the BatchRunNo, instead of the indicator
        clsBatch.Launch m_sBatchNumber, CStr(dExecution), CLng(sBatchrunNo)
        
        'Set the status to waiting
        'SYS4368 Commented out the SetStatus as it sets it on the wrong BatchRunNumber,
        ' and launching the batch sets the status anyway
        'm_clsBatchSchedule.SetStatus CLng(nStatusRunning)
        
        'Set the batchretryindicator to 1, to indicate the record has been rerun
        m_clsBatchSchedule.SetRetryBatchIndicator 1
        m_clsTableAccess.Update
        
        'SYS4410 The data must be reread from the database before calling PopulateListView
        'as launching the batch above will have updated the table
        'BM0425 PopulateListView retrieves the data already so this isn't required
        'm_clsBatchSchedule.GetBatchScheduleTableData m_sBatchNumber
        'SYS4410 End
        
        'BM0425 Reselect the previously selected item
        nSelectedItem = lvBatchInstance.SelectedItem.Index
        PopulateListView
        lvBatchInstance.ListItems.Item(nSelectedItem).Selected = True
        'BM0425 End

    Else
        g_clsErrorHandling.RaiseError errGeneralError, "This Schedule has already been rerun"
    End If
    
    EndWaitCursor   'BM0425
    'clsTableAccess.CancelFilter
    
    Exit Sub
Failed:
    EndWaitCursor   'BM0425
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

'BM0425
Private Sub Form_Resize()
    Const clBorder    As Long = 240
    Const clMinWidth  As Long = 6840
    Const clMinHeight As Long = 2840
    Dim lButtonHeight As Long

' Ignore resizing errors
On Error Resume Next
    
    If Me.Width < clMinWidth Then
        Me.Width = clMinWidth
    End If

    If Me.Height < clMinHeight Then
        Me.Height = clMinHeight
    End If
    
    lvBatchInstance.Width = Me.Width - (clBorder * 2)
    lvBatchInstance.Height = Me.Height - (clBorder * 8)
    cmdOK.Left = Me.Width - cmdOK.Width - clBorder - 100
    lButtonHeight = lvBatchInstance.Top + lvBatchInstance.Height + clBorder
    cmdRunFailedBatchJob.Top = lButtonHeight
    cmdViewErrors.Top = lButtonHeight
    cmdCancel.Top = lButtonHeight
    cmdOK.Top = lButtonHeight
    
End Sub
'BM0425 End

Private Sub lvBatchInstance_ItemClick(ByVal Item As MSComctlLib.IListItem)

    On Error GoTo Failed
    
    SetButtonState
        
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION

End Sub

Private Sub SetButtonState()
    On Error GoTo Failed
    
    Dim sBatchrunNo As String
    Dim sSQL As String
    Dim sBatchStatus As String
    
    lvBatchInstance.GetSelectedColumn m_nBatchRunNoColumn, sBatchrunNo
    ' Check if the status is failed - if not - there are no errors to view!
    sSQL = "BatchRunNumber = " & sBatchrunNo
    'Apply filter so the current record is the one selected
    m_clsTableAccess.ApplyFilter sSQL
    sBatchStatus = m_clsBatchSchedule.GetStatusValidationType
    Select Case sBatchStatus
        Case "W"
            cmdCancel.Enabled = True
            cmdRunFailedBatchJob.Enabled = False
        Case "R"
            cmdCancel.Enabled = True
            cmdRunFailedBatchJob.Enabled = False
        Case "CA"
            cmdCancel.Enabled = False
            cmdRunFailedBatchJob.Enabled = False
        Case "C"
            cmdCancel.Enabled = False
            cmdRunFailedBatchJob.Enabled = False
        Case "CWE"
            cmdCancel.Enabled = False
            cmdRunFailedBatchJob.Enabled = True
        'SYS4424 Button states were not being set for Initialisation Errors
        Case "IE"
            cmdCancel.Enabled = True
            cmdRunFailedBatchJob.Enabled = True
        'SYS4424 End
    End Select
    
    'BM0425 View Errors is now always enabled (SYS4557) so there is no point in checking the number of errors
    ''SYS4557 Revision to vieing of eror screen
    '
    ''SYS4424 The View Errors button should only be enabled if there are errors in the batch audit table
    'If m_clsBatchSchedule.GetNumberOfAuditErrors > 0 Then
    '    cmdViewErrors.Enabled = True
    'Else
    '    'cmdViewErrors.Enabled = False
    'End If
    '
    ''END SYS4557
    'BM0425
    
    'Batches that have already been rerun cannot be run again, so disable the button
    If m_clsBatchSchedule.GetBatchRetryIndicator <> 0 Then
        cmdRunFailedBatchJob.Enabled = False
    End If
    'SYS4424 End
    m_clsTableAccess.CancelFilter

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION

End Sub

