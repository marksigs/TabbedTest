VERSION 5.00
Object = "{24A7A369-0EDB-4924-817E-D81FE8E09890}#1.0#0"; "MSGOCX.ocx"
Begin VB.Form frmEditStages 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add/Edit Stages"
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9540
   Icon            =   "frmEditStages.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   9540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdDeleteTask 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   2520
      TabIndex        =   8
      Top             =   5460
      Width           =   1095
   End
   Begin VB.CommandButton cmdCreateTask 
      Caption         =   "&CreateTask"
      Height          =   375
      Left            =   4020
      TabIndex        =   9
      Top             =   5460
      Width           =   1095
   End
   Begin VB.CommandButton cmdAddTask 
      Caption         =   "&Add"
      Height          =   375
      Left            =   180
      TabIndex        =   6
      Top             =   5460
      Width           =   1095
   End
   Begin VB.CommandButton cmdEditTask 
      Caption         =   "&Edit"
      Height          =   375
      Left            =   1380
      TabIndex        =   7
      Top             =   5460
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   7080
      TabIndex        =   10
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   8280
      TabIndex        =   11
      Top             =   5880
      Width           =   1095
   End
   Begin MSGOCX.MSGListView lvTaskDetails 
      Height          =   2835
      Left            =   180
      TabIndex        =   5
      Top             =   2520
      Width           =   9195
      _ExtentX        =   16219
      _ExtentY        =   5001
      Sorted          =   -1  'True
      AllowColumnReorder=   0   'False
   End
   Begin VB.Frame Frame1 
      Height          =   2235
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   9195
      Begin MSGOCX.MSGComboBox cboUserAuthority 
         Height          =   315
         Left            =   1740
         TabIndex        =   3
         Top             =   1320
         Width           =   2715
         _ExtentX        =   4789
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
      Begin VB.CheckBox chkExceptionStage 
         Alignment       =   1  'Right Justify
         Caption         =   "Exception Stage?"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   1740
         Width           =   1815
      End
      Begin MSGOCX.MSGEditBox txtStage 
         Height          =   315
         Index           =   0
         Left            =   1740
         TabIndex        =   0
         Top             =   240
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   556
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
         MaxLength       =   20
      End
      Begin MSGOCX.MSGEditBox txtStage 
         Height          =   315
         Index           =   1
         Left            =   1740
         TabIndex        =   1
         Top             =   600
         Width           =   7035
         _ExtentX        =   12409
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
      Begin MSGOCX.MSGTextMulti txtRuleRef 
         Height          =   315
         Left            =   1740
         TabIndex        =   2
         Top             =   960
         Width           =   7035
         _ExtentX        =   12409
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
         MaxLength       =   100
      End
      Begin VB.Label Label6 
         Caption         =   "User Authority Level"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   1380
         Width           =   1515
      End
      Begin VB.Label Label1 
         Caption         =   "Stage ID"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   300
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Stage  Name"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   660
         Width           =   1335
      End
      Begin VB.Label Label10 
         Caption         =   "Rule Reference"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   1020
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmEditStages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form          :   frmEditStages
' Description   :   Form which allows the adding/editing of stages for Task Management

' Change history
' Prog      Date        Description
' DJP       09/11/00    Created for Phase 2 Task Management
' DJP       17/12/01    SYS2831 Client variants, GetListViewKeys moved to formProcessing.
' STB       10/04/02    SYS3878 Allow StageTasks removed from a stage to be re-added.
' STB       08/07/2002  SYS4529 BorderStyle set to 'Fixed Dialog'.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' BMIDS Specific
' CL        08/10/02    BMIDS00604 Modification to ValidateScreenData
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' MARS Specific
' GHun      10/03/2006  MAR1300 Add Process Authority Level column
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Epsom specific
' PB        24/04/2006  EP367   Changes merged into Epsom Supervisor
' TW        19/02/2007  EP2_1348 - Error in certain circumstances on MSGListView when AllowColumnReorder set to true
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

' Private data
Private m_sStageID As String
Private m_bIsEdit As Boolean
Private m_clsStageTable As StageTable
Private m_clsStageTaskTable As StageTaskTable
Private m_ReturnCode As MSGReturnCode
Private m_colKeys As Collection
Private m_nCurrentIndex As Long

' Private constants
Private Const STAGE_ID = 0
Private Const STAGE_NAME = 1
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SetIsEdit
' Description   :   Sets whether or not this form is in edit or add mode. Defaults to edit
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetIsEdit(Optional bIsEdit As Boolean = True)
    m_bIsEdit = bIsEdit
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   cmdEditTask_Click
' Description   :   Called when the user presses the Edit Task button
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdEditTask_Click()
    On Error GoTo Failed
    
    LoadStageTask True
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   cmdAddTask_Click
' Description   :   Called when the user presses the Add Task button
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdAddTask_Click()
    On Error GoTo Failed
    
    LoadStageTask
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
Private Sub SetCurrentTask()
    On Error GoTo Failed
    
    Dim bFound As Boolean
    Dim bFieldSuccess As Boolean
    Dim nCurrentRecord As Long
    Dim lstItem As ListItem
    Dim rs As ADODB.Recordset
    Dim colKeyValues As Collection
    Dim colKeys As Collection
    Dim sField As String
    Dim sValue As Variant
    Dim nFieldCount As Integer
    Dim nThisField As Integer
    
    Set lstItem = lvTaskDetails.SelectedItem
    
    g_clsFormProcessing.GetListViewKeys Me.lvTaskDetails, colKeyValues
    
    If Not colKeyValues Is Nothing Then
        Set colKeys = TableAccess(m_clsStageTaskTable).GetKeyMatchFields()
        
        If TableAccess(m_clsStageTaskTable).RecordCount() > 0 Then
            Set rs = TableAccess(m_clsStageTaskTable).GetRecordSet()
            
            rs.MoveFirst
            bFound = False
            nFieldCount = colKeys.Count
            
            Do While Not rs.EOF And Not bFound
                
                bFieldSuccess = True
                nThisField = 1
                
                Do While nThisField <= nFieldCount And bFieldSuccess
                    sField = colKeys(nThisField)
                    sValue = rs(sField).Value
                
                    If Not IsNull(sValue) Then
                        If sValue <> colKeyValues(nThisField) Then
                            bFieldSuccess = False
                        End If
                    Else
                        bFieldSuccess = False
                    End If
                    
                    nThisField = nThisField + 1
                Loop
                
                If bFieldSuccess Then
                    bFound = True
                Else
                    rs.MoveNext
                End If
                

            Loop
            
            If Not bFound Then
                g_clsErrorHandling.RaiseError errGeneralError, "Unable to locate Stage Task record"
            End If
            
        End If
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub LoadStageTask(Optional bIsEdit As Boolean = False)
    
    Dim sTaskID As String
    Dim enumReturn As MSGReturnCode
    Dim colExistingTasks As Collection
    
    On Error GoTo Failed
    
    If bIsEdit Then
        SetCurrentTask
        g_clsMainSupport.SetSelectedItem lvTaskDetails, m_nCurrentIndex, GET_SELECTION
    End If
    
    frmEditStageTask.SetIsEdit bIsEdit
    frmEditStageTask.SetTableAccess m_clsStageTaskTable
    
    ' Get list of existing tasks
    GetExistingTasks colExistingTasks, bIsEdit
    frmEditStageTask.SetExistingTasks colExistingTasks
    frmEditStageTask.SetStageID m_sStageID
    
    frmEditStageTask.Show vbModal, Me
        
    enumReturn = frmEditStageTask.GetReturnCode()
    
    If enumReturn = MSGSuccess Then
        'STB: SYS3878 - If we've just added a StageTask, ensure it isn't a duplicate
        'with one which has previously had its delete flag set to 1 by deleting
        'any other record(s) with the same TaskID.
        If bIsEdit = False Then
            sTaskID = m_clsStageTaskTable.GetTaskID
            m_clsStageTaskTable.CropDeletedTasks m_sStageID, sTaskID
        End If
        'STB : SYS3878 - End.
    
        ' Populate the Stage task listview
        PopulateStageTasks
    Else
    
    End If
    
    Unload frmEditStageTask
    
    If m_bIsEdit Then
        g_clsMainSupport.SetSelectedItem lvTaskDetails, m_nCurrentIndex, SET_SELECTION
    End If
    
    SetButtonState
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub GetExistingTasks(colExistingTasks As Collection, Optional bDoNotIncludeCurrent As Boolean = True)
    On Error GoTo Failed
    Dim rsTasks As ADODB.Recordset
    Dim rsTasksClone As ADODB.Recordset
    Dim sTaskID As String
    Set colExistingTasks = New Collection
    
    If TableAccess(m_clsStageTaskTable).RecordCount() > 0 Then
        ' If we're editing, get the current task ID as we don't want to include it in the list
        If bDoNotIncludeCurrent Then
            Dim sCurrentTaskID As String
            
            sCurrentTaskID = m_clsStageTaskTable.GetTaskID()
        
        End If
        
        Set rsTasks = TableAccess(m_clsStageTaskTable).GetRecordSet()
        Set rsTasksClone = rsTasks.Clone()
        
        If Not rsTasksClone Is Nothing Then
            TableAccess(m_clsStageTaskTable).SetRecordSet rsTasksClone
            
            rsTasksClone.MoveFirst
            Do While Not rsTasksClone.EOF
                sTaskID = m_clsStageTaskTable.GetTaskID()
                
                If sCurrentTaskID <> sTaskID Then
                    colExistingTasks.Add sTaskID
                End If
                
                rsTasksClone.MoveNext
            Loop
            
            rsTasksClone.Close
            Set rsTasksClone = Nothing
            TableAccess(m_clsStageTaskTable).SetRecordSet rsTasks
        
        End If
        
    End If
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub PopulateStageTasks()
    On Error GoTo Failed
        
    g_clsFormProcessing.PopulateFromRecordset lvTaskDetails, m_clsStageTaskTable
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   cmdCreateTask_Click
' Description   :   Called when the user presses the Create Task button
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdCreateTask_Click()
    On Error GoTo Failed
    Dim enumReturn As MSGReturnCode
    
    frmEditTasks.SetIsEdit False
    frmEditTasks.Show vbModal, Me
    
    enumReturn = frmEditTasks.GetReturnCode()
        
    ' Any post add/edit processing?
    
    Unload frmEditTasks
        
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   cmdDeleteTask_Click
' Description   :   Called when the user presses the Delete Task button
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdDeleteTask_Click()
    On Error GoTo Failed
    
    DeleteCurrentTask
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
Private Sub DeleteCurrentTask()
    On Error GoTo Failed
    
    ' From the recordset
    SetCurrentTask
    m_clsStageTaskTable.SetDeleteFlag True
    
    ' From the listview
    Dim nCurrentRecord As Long
    Dim lstItems As ListItems
    Dim lstSelectedItem As ListItem
    Dim bEnable As Boolean
    Dim nCount As Long
    
    Set lstSelectedItem = lvTaskDetails.SelectedItem
    Set lstItems = lvTaskDetails.ListItems
    
    If Not lstSelectedItem Is Nothing Then
        nCurrentRecord = lstSelectedItem.Index
        lstItems.Remove nCurrentRecord
    End If
    
    bEnable = True
    nCount = lstItems.Count
    
    If nCurrentRecord > nCount And nCount > 0 Then
        nCurrentRecord = nCount
    ElseIf nCount = 0 Then
        bEnable = False
    End If
    
    If bEnable Then
        Set lvTaskDetails.SelectedItem = lstItems(nCurrentRecord)
    End If
    
    SetButtonState bEnable
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   Form_Load
' Description   :   VB Calls this method when the form is loaded
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
    On Error GoTo Failed

    Initialise

    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
    g_clsFormProcessing.CancelForm Me
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   Initialise
' Description   :   Funciton which performs all screen initialisation. Populate all combos, read
'                   data from the database if in add mode, create a new record if in add mode.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Initialise()
    On Error GoTo Failed
    m_ReturnCode = MSGFailure
    
    Set m_clsStageTable = New StageTable
    Set m_clsStageTaskTable = New StageTaskTable
    m_sStageID = ""
    
    PopulateCombos
        
    ' Set the listview headers
    SetListViewHeaders
    
    ' Edit mode?
    If m_bIsEdit Then
        SetEditState
    Else
        SetAddState
    End If

    ' Enable buttons
    SetButtonState

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SetButtonState
' Description   :   Sets the add/edit/add task/create tasks buttons depending on whether we are in edit
'                   or add mode.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetButtonState(Optional bEnable As Boolean = True)
    On Error GoTo Failed
    Dim bEnableEdit As Boolean
    
    cmdAddTask.Enabled = True
    cmdCreateTask.Enabled = True
    
    Dim lstSelection As ListItem
    Set lstSelection = Me.lvTaskDetails.SelectedItem
    
    If Not lstSelection Is Nothing Then
        bEnableEdit = True And bEnable
    Else
        bEnableEdit = False
    End If
    
    cmdEditTask.Enabled = bEnableEdit
    cmdDeleteTask.Enabled = bEnableEdit
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SetListViewHeaders
' Description   :   Sets the field header titles on the task listview
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetListViewHeaders()
    On Error GoTo Failed
    Dim headers As New Collection
    Dim lvHeaders As listViewAccess
    
    lvHeaders.nWidth = 20
    lvHeaders.sName = "Task ID"
    headers.Add lvHeaders
    
    lvHeaders.nWidth = 9
    lvHeaders.sName = "Default"
    headers.Add lvHeaders
    
    lvHeaders.nWidth = 9
    lvHeaders.sName = "Additional"
    headers.Add lvHeaders
    
    lvHeaders.nWidth = 9
    lvHeaders.sName = "Mandatory"
    headers.Add lvHeaders
    
    lvHeaders.nWidth = 9
    lvHeaders.sName = "Carry Forward"
    headers.Add lvHeaders
    
    lvHeaders.nWidth = 15
    lvHeaders.sName = "Trigger next stage?"
    headers.Add lvHeaders
    
    lvHeaders.nWidth = 15
    lvHeaders.sName = "Create Authority Level"
    headers.Add lvHeaders
    
    'MAR1300 GHun
    lvHeaders.nWidth = 15
    lvHeaders.sName = "Process Authority Level"
    headers.Add lvHeaders
    'MAR1300 End
    
    lvHeaders.nWidth = 10
    lvHeaders.sName = "User ID"
    headers.Add lvHeaders
    
    lvHeaders.nWidth = 10
    lvHeaders.sName = "Unit ID"
    headers.Add lvHeaders
    
    lvTaskDetails.AddHeadings headers
    lvTaskDetails.LoadColumnDetails TypeName(Me)
    
    lvTaskDetails.HideSelection = False
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   PopulateCombos
' Description   :   Populates combo boxes with their values. The come from a combo
'                   group (PopulateCombo)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PopulateCombos()
    On Error GoTo Failed

    ' User Authority Level
    g_clsFormProcessing.PopulateCombo "UserRole", cboUserAuthority

    Exit Sub
Failed:

    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   cmdCancel_Click
' Description   :   Called when the user cliks the Cancel button. All we do is hide the form, which
'                   passes control back to the caller, which checks the status of the form, and closes
'                   it.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdCancel_Click()
    On Error GoTo Failed

    Hide

    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   DoOKProcessing
' Description   :   Common function to be used by Another and OK. Validates the data on the screen
'                   and saves all screen data to the database. Also records the change just made
'                   using SaveChangeRequest
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub DoOKProcessing()
    On Error GoTo Failed

    ValidateScreenData
    SaveScreenData
    SaveChangeRequest

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   ValidateScreenData
' Description   :   Perform all screen validation.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ValidateScreenData()
    On Error GoTo Failed

    Dim vStages As Variant
    
    g_clsFormProcessing.DoMandatoryProcessing Me

    'Check that there is at least one task entered.
    vStages = lvTaskDetails.ListItems.Count
    'BMIDS00604 CL 8/10/02
    If vStages <> 0 Then
        'Check the record doesn't exist already, if doing an add
        If Not m_bIsEdit Then
            CheckStageExists
        End If
    Else
        'Raise an error as at least one task must be selected.
        g_clsErrorHandling.RaiseError errGeneralError, "At least one task must be included on a stage record"
    End If
    'BMIDS00604 CL 8/10/02 END
        
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub CheckStageExists()
    On Error GoTo Failed
    Dim bSuccess As Boolean
    Dim colMatchValues As Collection
    
    Set colMatchValues = New Collection
    colMatchValues.Add txtStage(STAGE_ID).Text
    
    bSuccess = Not TableAccess(m_clsStageTable).DoesRecordExist(colMatchValues)
    
    If Not bSuccess Then
        txtStage(STAGE_ID).SetFocus
        g_clsErrorHandling.RaiseError errGeneralError, "Stage already exists - please enter a unique stage"
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   cmdOK_Click
' Description   :   Called when the user clicks the OK button. If DoOKProcessing succeeds (i.e.,
'                   screen is validated and data saved), set the form return code to true and
'                   hide the form thus returning control back to the caller
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOK_Click()
    On Error GoTo Failed

    DoOKProcessing

    ' If we get here, all is ok
    SetReturnCode

    Hide

    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SetReturnCode
' Description   :   Sets the return code for the form for any calling method to check. Defaults
'                   to MSG_SUCCESS
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetReturnCode(Optional enumReturn As MSGReturnCode = MSGSuccess)
    m_ReturnCode = enumReturn
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   GetReturnCode
' Description   :   Returns the return code for this form, either MSG_SUCCESS or MSG_FAILURE
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetReturnCode() As MSGReturnCode
    GetReturnCode = m_ReturnCode
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SaveChangeRequest
' Description   :   Saves the fact that the current task has been created/edited for future promotion
'                   to other databases
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveChangeRequest()
    On Error GoTo Failed
    Dim colValues As Collection

    ' Get list of keys
    Set colValues = New Collection
    colValues.Add txtStage(STAGE_ID).Text

    TableAccess(m_clsStageTable).SetKeyMatchValues colValues

    ' Save this key
    g_clsHandleUpdates.SaveChangeRequest m_clsStageTable

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SaveScreenData
' Description   :   Saves all screen data to the database. All combo values actually store the
'                   combo ID and not the combo text.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveScreenData()
    On Error GoTo Failed
    Dim vval As Variant
        
    ' Stage ID
    m_sStageID = txtStage(STAGE_ID).Text

    m_clsStageTable.SetStageID m_sStageID

    ' Stage Name
    m_clsStageTable.SetStageName txtStage(STAGE_NAME).Text

    ' Rule Reference
    m_clsStageTable.SetRuleRef txtRuleRef.Text
    
    ' Authority
    g_clsFormProcessing.HandleComboExtra Me.cboUserAuthority, vval, GET_CONTROL_VALUE
    m_clsStageTable.SetAuthorityLevel vval
    
    ' Exception Stage
    g_clsFormProcessing.HandleCheckBox chkExceptionStage, vval, GET_CONTROL_VALUE
    m_clsStageTable.SetException vval

    If Not m_bIsEdit Then
        m_clsStageTable.SetActivityID Null
    End If
    
    m_clsStageTable.SetDeleteFlag

    ' Save the associated tasks
    SaveTasks

    TableAccess(m_clsStageTable).Update
    TableAccess(m_clsStageTaskTable).Update

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub SaveTasks()
    On Error GoTo Failed
    Dim sStageID As String
    Dim colUpdateValues As Collection
    
    Set colUpdateValues = New Collection
    colUpdateValues.Add m_sStageID
        
    BandedTable(m_clsStageTaskTable).SetUpdateValues colUpdateValues
    BandedTable(m_clsStageTaskTable).SetUpdateSets
    BandedTable(m_clsStageTaskTable).DoUpdateSets
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   PopulateScreenFields
' Description   :   Sets all screen fields to the values stored on the database
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PopulateScreenFields()
    On Error GoTo Failed

    Dim vval As Variant

    ' Stage ID
    m_sStageID = m_clsStageTable.GetStageID()
    txtStage(STAGE_ID).Text = m_sStageID

    ' Stage Name
    
    txtStage(STAGE_NAME).Text = m_clsStageTable.GetStageName()

    ' Rule Reference
    txtRuleRef.Text = m_clsStageTable.GetRuleRef()

    ' Authority
    vval = m_clsStageTable.GetAuthorityLevel()
    g_clsFormProcessing.HandleComboExtra Me.cboUserAuthority, vval, SET_CONTROL_VALUE
    
    ' Exception Stage
    vval = m_clsStageTable.GetException()
    g_clsFormProcessing.HandleCheckBox chkExceptionStage, vval, SET_CONTROL_VALUE

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SetAddState
' Description   :   Called on startup to do any Add specific processing
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetAddState()
    On Error GoTo Failed

    g_clsFormProcessing.CreateNewRecord m_clsStageTable
    TableAccess(m_clsStageTaskTable).GetTableData POPULATE_EMPTY

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SetEditState
' Description   :   Called on startup to do any Edit specific processing - will need to read the
'                   record required from the database via the task table class
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetEditState()
    On Error GoTo Failed

    TableAccess(m_clsStageTable).SetKeyMatchValues m_colKeys

    ' Get the data from the database
    TableAccess(m_clsStageTable).GetTableData

    ' Validate we have the record
    If TableAccess(m_clsStageTable).RecordCount() = 0 Then
        g_clsErrorHandling.RaiseError errGeneralError, "Edit Tasks - Unable to locate task"
    End If

    ' If we get here, we have the data we need
    PopulateScreenFields

    ' Populate the listview with the list of attached tasks
    PopulateTasks
    
    txtStage(STAGE_ID).Enabled = False

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   PopulateTasks
' Description   :   Populates the listview with all tasks attached to this stage
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PopulateTasks()
    On Error GoTo Failed
    Dim sStageNumber As String
    
    If Len(m_sStageID) = 0 Then
        g_clsErrorHandling.RaiseError errKeysEmpty, TypeName(Me) & " PopulateTasks"
    End If
    
    ' Get the tasks
    m_clsStageTaskTable.GetTasksForStage m_sStageID
    
    ' Populate the listview from the StageTask table class
    g_clsFormProcessing.PopulateFromRecordset lvTaskDetails, m_clsStageTaskTable
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SetKeys
' Description   :   Called from a calling method to set the key values used by this form to locate
'                   a record from the database when this form is initialsed
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetKeys(colKeys As Collection)
    On Error GoTo Failed
    Dim sFunctionName As String

    sFunctionName = ", SetKeys"

    If Not colKeys Is Nothing Then
        If colKeys.Count > 0 Then
            Set m_colKeys = colKeys
        Else
            g_clsErrorHandling.RaiseError errKeysEmpty, TypeName(Me) & sFunctionName
        End If
    Else
        g_clsErrorHandling.RaiseError errKeysEmpty, TypeName(Me) & sFunctionName
    End If

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub Form_Initialize()
    m_bIsEdit = False
    m_nCurrentIndex = -1
End Sub
Private Sub lvTaskDetails_DblClick()
    On Error GoTo Failed
    
    LoadStageTask True
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
Private Sub lvTaskDetails_DeletePressed()
    On Error GoTo Failed
    
    DeleteCurrentTask
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   lvTaskDetails_ItemClick
' Description   :   Called when an itme on the listview is clicked. Need to enable/disable all
'                   relevant buttons
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub lvTaskDetails_ItemClick(ByVal Item As MSComctlLib.IListItem)
    On Error GoTo Failed
    
    SetButtonState
        
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
