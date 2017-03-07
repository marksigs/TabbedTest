VERSION 5.00
Object = "{70BDCEF6-8EE8-48A3-BD88-91622E167DCB}#1.1#0"; "MSGOCX.ocx"
Begin VB.Form frmEditStageTask 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add/Edit Stage Task"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8175
   Icon            =   "frmEditStageTask.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   8175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7935
      Begin VB.CheckBox chkMandatory 
         Alignment       =   1  'Right Justify
         Caption         =   "Mandatory for this stage?"
         Height          =   255
         Left            =   210
         TabIndex        =   12
         Top             =   1440
         Width           =   2560
      End
      Begin VB.CheckBox chkAdditionalTask 
         Alignment       =   1  'Right Justify
         Caption         =   "Additional Task"
         Height          =   255
         Left            =   210
         TabIndex        =   11
         Top             =   1080
         Width           =   2560
      End
      Begin VB.CheckBox chkDefaultTask 
         Alignment       =   1  'Right Justify
         Caption         =   "Default Task"
         Height          =   255
         Left            =   210
         TabIndex        =   10
         Top             =   720
         Width           =   2560
      End
      Begin VB.CheckBox chkCarriedForward 
         Alignment       =   1  'Right Justify
         Caption         =   "Can be carried forward?"
         Height          =   255
         Left            =   210
         TabIndex        =   9
         Top             =   1800
         Width           =   2560
      End
      Begin VB.CommandButton cmdAnother 
         Caption         =   "&Another"
         Height          =   375
         Left            =   6480
         TabIndex        =   8
         Top             =   4080
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         CausesValidation=   0   'False
         Height          =   375
         Left            =   1440
         TabIndex        =   7
         Top             =   4080
         Width           =   1095
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   4080
         Width           =   1095
      End
      Begin VB.CheckBox chkTriggerNextStage 
         Alignment       =   1  'Right Justify
         Caption         =   "Triggers next stage?"
         Height          =   255
         Left            =   210
         TabIndex        =   1
         Top             =   2160
         Width           =   2560
      End
      Begin MSGOCX.MSGComboBox cboUnitID 
         Height          =   315
         Left            =   2580
         TabIndex        =   2
         Top             =   3180
         Width           =   2895
         _ExtentX        =   5106
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
      Begin MSGOCX.MSGComboBox cboAuthorityLevel 
         Height          =   315
         Left            =   2580
         TabIndex        =   3
         Top             =   2460
         Width           =   2895
         _ExtentX        =   5106
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
      Begin MSGOCX.MSGDataCombo cboUserID 
         Height          =   315
         Left            =   2580
         TabIndex        =   4
         Top             =   3540
         Width           =   2895
         _ExtentX        =   5106
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
      End
      Begin MSGOCX.MSGComboBox cboTask 
         Height          =   315
         Left            =   2580
         TabIndex        =   5
         Top             =   285
         Width           =   5055
         _ExtentX        =   8916
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
         Mandatory       =   -1  'True
         Text            =   ""
      End
      Begin MSGOCX.MSGComboBox cboProcessAuthorityLevel 
         Height          =   315
         Left            =   2580
         TabIndex        =   17
         Top             =   2820
         Width           =   2895
         _ExtentX        =   5106
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
      Begin VB.Label lblProcessAuthority 
         Caption         =   "Process Task Authority Level"
         Height          =   195
         Left            =   240
         TabIndex        =   18
         Top             =   2880
         Width           =   2175
      End
      Begin VB.Label lblCreateAuthority 
         Caption         =   "Create Task Authority Level"
         Height          =   195
         Left            =   240
         TabIndex        =   16
         Top             =   2520
         Width           =   2295
      End
      Begin VB.Label Label6 
         Caption         =   "Task"
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "User ID"
         Height          =   195
         Left            =   240
         TabIndex        =   14
         Top             =   3600
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Unit ID"
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   3240
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmEditStageTask"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form          :   frmEditStageTask
' Description   :   Form which allows the adding/editing of Stage Tasks for Task Management

' Change history
' Prog      Date        Description
' DJP       09/11/00    Created for Phase 2 Task Management
' STB       06/12/01    SYS1942 - 'Another' button / transactions.
' STB       10/04/02    SYS3878 - Allow removed StageTasks to be re-added to a stage.
' STB       08/07/2002  SYS4529 BorderStyle set to 'Fixed Dialog'.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' BMIDS Change History:
'
' Prog      Date        Description
' DB        14/11/2002  BMIDS00877 Added check to ensure a mandatory task cannot be carried forward
' HMA       19/10/2004  BMIDS603   Clear User when the Unit is deselected.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' MARS Change History:
'
' Prog      Date        Description
' GHun      07/03/2006  MAR1300 Added ProcessTaskAuthorityLevel
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

' Private data
Private m_bIsEdit As Boolean
Private m_clsStageTaskTable As StageTaskTable
Private m_ReturnCode As MSGReturnCode
Private m_colExistingTasks  As Collection
Private m_sStageID As String
Private m_sTaskID As String

'Used to indicate if cmdAnother has been used, this will store the number of
'records added in one 'session'.
Private m_iRecordsAdded As Integer

' Private constants
Private Const TASK_ID = 0
Private Const TASK_NAME = 1
Private Const CHASING_PERIOD_DAYS = 2
Private Const CHASING_PERIOD_HOURS = 3
Private Const INPUT_PROCESS = 4
Private Const OUTPUT_DOCUMENT = 5
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SetIsEdit
' Description   :   Sets whether or not this form is in edit or add mode. Defaults to edit
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetStageID(sStageID As String)
    m_sStageID = sStageID
End Sub
Public Sub SetIsEdit(Optional bIsEdit As Boolean = True)
    m_bIsEdit = bIsEdit
End Sub
Private Sub cboUnitID_Click()
    On Error GoTo Failed
    Dim sUnitID As String
    
    sUnitID = cboUnitID.SelText
    
    'BMIDS603  IF the Unit is deselected, clear the user.
    If Len(sUnitID) > 0 Then
        PopulateUserID sUnitID
    Else
        cboUserID.SelText = ""
    End If
    
    SetCombosState
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub
Public Sub SetExistingTasks(colExistingTasks As Collection)
    Set m_colExistingTasks = colExistingTasks
End Sub

Private Sub cmdAnother_Click()
    On Error GoTo Failed
    
    'Trip this flag to true.
    m_iRecordsAdded = m_iRecordsAdded + 1
    
    DoOKProcessing
    
    g_clsFormProcessing.ClearScreenFields Me
    
    SetAddState
    
    m_colExistingTasks.Add m_sTaskID
    PopulateTask
    
    cboTask.SetFocus
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   Form_Load
' Description   :   VB Calls this method when the form is loaded
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
    On Error GoTo Failed
    m_sTaskID = ""
    
    'Set the number of records modified to one.
    m_iRecordsAdded = 1
    
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

    ' Task table contains all database specific code
'    Set m_clsStageTaskTable = New StageTaskTable
    m_ReturnCode = MSGFailure

    PopulateCombos

    ' Edit mode?
    If m_bIsEdit Then
        SetEditState
    Else
        SetAddState
    End If
 
    SetCombosState

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   PopulateCombos
' Description   :   Populates combo boxes with their values. The values may come from a combo
'                   group (PopulateCombo), or from other tables (PopulateChasingTask)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PopulateCombos()
    On Error GoTo Failed

    ' Authority level
    g_clsFormProcessing.PopulateCombo "UserRole", cboAuthorityLevel
    g_clsFormProcessing.PopulateCombo "UserRole", cboProcessAuthorityLevel  'MAR1300

    ' Unit ID
    PopulateUnitID
    
    ' Task
    PopulateTask

    Exit Sub
Failed:

    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SetCombosState
' Description   :   Sets the state of the unit and user combos. If an entry exists in the unit
'                   combo, enable the user combo and make it mandatory. If no entry exists in
'                   the unit combo, disable the user combo
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetCombosState()
    On Error GoTo Failed
    Dim sUnitID As String
    Dim bEnableUser As Boolean
    
    bEnableUser = False
    
    sUnitID = cboUnitID.SelText
    
    If Len(sUnitID) > 0 Then
        bEnableUser = True
    End If
        
    cboUserID.Enabled = bEnableUser
    cboUserID.Mandatory = bEnableUser
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   PopulateUserID
' Description   :   Populates the UserId combobox with a list of active user id's.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PopulateUserID(Optional sUnit As String = "")
    On Error GoTo Failed
    Dim clsUser As OmigaUserTable
    Dim rsUser As ADODB.Recordset
    Dim sListField As String
    
    Set clsUser = New OmigaUserTable
    cboUserID.SelText = ""
    
    If Len(sUnit) > 0 Then
        clsUser.GetUsersFromUnit sUnit
    Else
        clsUser.GetActiveUsers
    End If
    
    sListField = clsUser.GetComboField()
    
    Set rsUser = TableAccess(clsUser).GetRecordSet()
    Set cboUserID.RowSource = rsUser
    cboUserID.ListField = sListField
    
    If rsUser.RecordCount = 0 Then
        cboUserID.Enabled = False
        g_clsErrorHandling.RaiseError errGeneralError, "No Users exist for the selected Unit"
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   Populate UnitID
' Description   :   Populates the UnitID combobox with a list of active Units (but only units
'                   that have users attached to them)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PopulateUnitID()
    On Error GoTo Failed
    Dim clsUnit As UnitTable
    Dim sListField As String
    Dim colField As Collection
    
    Set clsUnit = New UnitTable
    Set colField = New Collection
    
    clsUnit.GetActiveUnits
    
    sListField = clsUnit.GetComboField()
    colField.Add sListField
    
    g_clsFormProcessing.PopulateComboFromTable cboUnitID, clsUnit, colField
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   PopulateTask
' Description   :   Populates the Task combo box with a list of all Tasks read from the
'                   database (the Task table)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PopulateTask()
    On Error GoTo Failed
    Dim colFields As Collection
    Dim clsTaskTable As TaskTable
    
    ' Create the Task table access class
    Set clsTaskTable = New TaskTable
    
    'STB: SYS3878 - Only exclude items in the collection.
    clsTaskTable.GetTasks m_colExistingTasks
    'STB: SYS3878 - End.
    
    Set colFields = clsTaskTable.GetComboFields()
    g_clsFormProcessing.PopulateComboFromTable cboTask, clsTaskTable, colFields
    
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
    
    Dim uResponse As VbMsgBoxResult
    
    On Error GoTo Failed

    'If the another button has been used, we should warn that more than one record
    'is about to be lost. Normally this would only rollback the last record added,
    'but this form is a popup which runs in a wider transaction.
    If m_iRecordsAdded > 1 Then
        uResponse = MsgBox("You have added more than one record. Proceeding will cancel all records added. Do you wish to proceed?", vbYesNo Or vbExclamation, "Loose Changes?")
    End If
    
    'If editing, only one record has been added or the user is willing to loose
    'multiple changes, then allow the cancel to proceed,
    If (m_iRecordsAdded < 2) Or (uResponse = vbYes) Then
        If Not m_bIsEdit Then
            CancelCurrentRecord
        End If
    
        Hide
    End If

    Exit Sub
    
Failed:
    g_clsErrorHandling.DisplayError
End Sub


Private Sub CancelCurrentRecord()
    On Error GoTo Failed
    Dim rsStageTask As ADODB.Recordset
    
    Set rsStageTask = TableAccess(m_clsStageTaskTable).GetRecordSet()

    If Not rsStageTask Is Nothing Then
        rsStageTask.CancelBatch adAffectCurrent
    End If

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
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
    'SaveChangeRequest

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
    g_clsFormProcessing.DoMandatoryProcessing Me
    
'    DoesStageTaskExist

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'Private Sub DoesStageTaskExist()
'    On Error GoTo Failed
'    Dim sTaskID As String
'    Dim colMatchValues As Collection
'    Dim bExists As Boolean
'
'    g_clsFormProcessing.HandleComboExtra cboTask, sTaskID, GET_CONTROL_VALUE
'    Set colMatchValues = New Collection
'
'    colMatchValues.Add sTaskID
'
'    bExists = TableAccess(m_clsStageTaskTable).DoesRecordExist(colMatchValues)
'
'    If bExists Then
'        g_clsErrorHandling.RaiseError errGeneralError, "Task exists for this stage"
'    End If
'
'    Exit Sub
'Failed:
'    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
'End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   cmdOK_Click
' Description   :   Called when the user clicks the OK button. If DoOKProcessing succeeds (i.e.,
'                   screen is validated and data saved), set the form return code to true and
'                   hide the form thus returning control back to the caller
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOK_Click()
    On Error GoTo Failed
    
    'DB BMIDS00877 - cannot be mandatory task and carried forward
    If chkMandatory.Value = 1 And chkCarriedForward.Value = 1 Then
        MsgBox "You cannot carry forward a mandatory task", vbExclamation
    Else
        DoOKProcessing
        ' If we get here, all is ok
        SetReturnCode
        Hide
    End If
    'DB END
    
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
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Function      :   SaveChangeRequest
'' Description   :   Saves the fact that the current task has been created/edited for future promotion
''                   to other databases
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Private Sub SaveChangeRequest()
'    On Error GoTo Failed
'    Dim colValues As Collection
'
'    ' Get list of keys
'    Set colValues = New Collection
'    colValues.Add Me.txtTask(TASK_ID).Text
'
'    TableAccess(m_clsTaskTable).SetKeyMatchValues colValues
'
'    ' Save this key
'    g_clsHandleUpdates.SaveChangeRequest m_clsTaskTable
'
'    Exit Sub
'Failed:
'    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
'End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SaveScreenData
' Description   :   Saves all screen data to the database. All combo values actually store the
'                   combo ID and not the combo text.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveScreenData()
    On Error GoTo Failed
    Dim sTaskID As String
    Dim vval As Variant

    ' Task
    g_clsFormProcessing.HandleComboExtra cboTask, vval, GET_CONTROL_VALUE
    m_clsStageTaskTable.SetTaskId vval
    m_sTaskID = CStr(vval)
    
    ' Default Task
    g_clsFormProcessing.HandleCheckBox chkDefaultTask, vval, GET_CONTROL_VALUE
    m_clsStageTaskTable.SetDefaultTask vval
    
    ' Additional Task
    g_clsFormProcessing.HandleCheckBox chkAdditionalTask, vval, GET_CONTROL_VALUE
    m_clsStageTaskTable.SetAdditionalTask vval
    
    ' Mandatory for this stage?
    g_clsFormProcessing.HandleCheckBox chkMandatory, vval, GET_CONTROL_VALUE
    m_clsStageTaskTable.SetMandatory vval
    
    ' Can be carried forward?
    g_clsFormProcessing.HandleCheckBox chkCarriedForward, vval, GET_CONTROL_VALUE
    m_clsStageTaskTable.SetCarriedForward vval
    
    ' Trigger Task
    g_clsFormProcessing.HandleCheckBox chkTriggerNextStage, vval, GET_CONTROL_VALUE
    m_clsStageTaskTable.SetTriggerTask vval
    
    ' Create Task Authority level
    g_clsFormProcessing.HandleComboExtra cboAuthorityLevel, vval, GET_CONTROL_VALUE
    m_clsStageTaskTable.SetAuthorityLevel vval
    
    'MAR1300 GHun
    ' Process Task Authority level
    g_clsFormProcessing.HandleComboExtra cboProcessAuthorityLevel, vval, GET_CONTROL_VALUE
    m_clsStageTaskTable.SetProcessAuthorityLevel vval
    'MAR1300 End
    
    ' User ID
    g_clsFormProcessing.HandleDataComboText cboUserID, vval, GET_CONTROL_VALUE
    m_clsStageTaskTable.SetUserID vval
    
    ' Unit ID
    g_clsFormProcessing.HandleComboText cboUnitID, vval, GET_CONTROL_VALUE
    m_clsStageTaskTable.SetUnitID vval
    
    m_clsStageTaskTable.SetDeleteFlag

    'TableAccess(m_clsStageTaskTable).Update

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

    ' Task
    vval = m_clsStageTaskTable.GetTaskId()
    g_clsFormProcessing.HandleComboExtra cboTask, vval, SET_CONTROL_VALUE

    ' Default Task
    vval = m_clsStageTaskTable.GetDefaultTask()
    g_clsFormProcessing.HandleCheckBox chkDefaultTask, vval, SET_CONTROL_VALUE
    
    ' Additional Task
    vval = m_clsStageTaskTable.GetAdditionalTask()
    g_clsFormProcessing.HandleCheckBox chkAdditionalTask, vval, SET_CONTROL_VALUE
    
    ' Mandatory for this stage?
    vval = m_clsStageTaskTable.GetMandatory()
    g_clsFormProcessing.HandleCheckBox chkMandatory, vval, SET_CONTROL_VALUE
    
    ' Can be carried forward?
    vval = m_clsStageTaskTable.GetCarriedForward()
    g_clsFormProcessing.HandleCheckBox chkCarriedForward, vval, SET_CONTROL_VALUE

    ' Trigger Task?
    vval = m_clsStageTaskTable.GetTriggerTask()
    g_clsFormProcessing.HandleCheckBox chkTriggerNextStage, vval, SET_CONTROL_VALUE

    ' Create Task Authority level
    vval = m_clsStageTaskTable.GetAuthorityLevel()
    g_clsFormProcessing.HandleComboExtra cboAuthorityLevel, vval, SET_CONTROL_VALUE
    
    'MAR1300 GHun
    ' Process Task Authority level
    vval = m_clsStageTaskTable.GetProcessAuthorityLevel()
    g_clsFormProcessing.HandleComboExtra cboProcessAuthorityLevel, vval, SET_CONTROL_VALUE
    'MAR1300 End
    
    ' Unit ID
    vval = m_clsStageTaskTable.GetUnitID()
    g_clsFormProcessing.HandleComboText cboUnitID, vval, SET_CONTROL_VALUE

    If Len(vval) > 0 Then
        PopulateUserID CStr(vval)
    
        ' User ID
        vval = m_clsStageTaskTable.GetUserID()
        g_clsFormProcessing.HandleDataComboText cboUserID, vval, SET_CONTROL_VALUE
    End If
    
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

    TableAccess(m_clsStageTaskTable).AddRow

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

    'TableAccess(m_clsTaskTable).SetKeyMatchValues m_colKeys

    ' Get the data from the database
    'TableAccess(m_clsStageTaskTable).GetTableData

    ' Validate we have the record
    If TableAccess(m_clsStageTaskTable).RecordCount() = 0 Then
        g_clsErrorHandling.RaiseError errGeneralError, "Edit Tasks - Unable to locate task"
    End If

    ' If we get here, we have the data we need
    PopulateScreenFields

    cmdAnother.Enabled = False

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SetTableAccess
' Description   :   Called from a calling method to set the table class to be used for database
'                   reads/writes
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetTableAccess(clsTableAccess As TableAccess)
    On Error GoTo Failed
    Dim sFunctionName As String

    sFunctionName = "SetTableAccess"

    If Not clsTableAccess Is Nothing Then
        Set m_clsStageTaskTable = clsTableAccess
    Else
        g_clsErrorHandling.RaiseError errRecordSetEmpty, " " & sFunctionName
    End If

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub Form_Initialize()
    m_bIsEdit = False
    Set m_colExistingTasks = Nothing
End Sub


' TRIGGERTASKIND STAGETASKTABLE
' Only one task on the stage can have this set to true.

