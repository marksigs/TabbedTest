VERSION 5.00
Object = "{61B85B2D-AE68-4B98-B8E9-B759DBAF18E8}#1.2#0"; "MSGOCX.ocx"
Begin VB.Form frmEditBusinessGroup 
   Caption         =   "Add/Edit Business Group"
   ClientHeight    =   6300
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7530
   Icon            =   "frmEditBusinessGroup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   7530
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   1380
      TabIndex        =   8
      Top             =   5820
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   60
      TabIndex        =   7
      Top             =   5820
      Width           =   1215
   End
   Begin MSGOCX.MSGComboBox cboBusinessArea 
      Height          =   315
      Left            =   1680
      TabIndex        =   1
      Top             =   600
      Width           =   5715
      _ExtentX        =   10081
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
   Begin MSGOCX.MSGTextMulti txtCompleteText 
      Height          =   495
      Left            =   1680
      TabIndex        =   2
      Top             =   1020
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   873
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
      MaxLength       =   50
   End
   Begin MSGOCX.MSGHorizontalSwapList swapTasks 
      Height          =   2535
      Left            =   60
      TabIndex        =   6
      Top             =   3300
      Width           =   7395
      _ExtentX        =   13044
      _ExtentY        =   4471
   End
   Begin MSGOCX.MSGEditBox txtGroupName 
      Height          =   315
      Left            =   1680
      TabIndex        =   0
      Top             =   180
      Width           =   5715
      _ExtentX        =   10081
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
   Begin MSGOCX.MSGTextMulti txtIncompleteText 
      Height          =   495
      Left            =   1680
      TabIndex        =   3
      Top             =   1620
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   873
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
      MaxLength       =   50
   End
   Begin MSGOCX.MSGTextMulti txtNotApplicableText 
      Height          =   495
      Left            =   1680
      TabIndex        =   4
      Top             =   2220
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   873
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
      MaxLength       =   50
   End
   Begin MSGOCX.MSGTextMulti txtNotStartedText 
      Height          =   495
      Left            =   1680
      TabIndex        =   5
      Top             =   2820
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   873
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
      MaxLength       =   50
   End
   Begin VB.Label lblNotStarted 
      Caption         =   "Not Started Text"
      Height          =   195
      Left            =   60
      TabIndex        =   14
      Top             =   2820
      Width           =   1335
   End
   Begin VB.Label lblNotApplicable 
      Caption         =   "Not Applicable Text"
      Height          =   375
      Left            =   60
      TabIndex        =   13
      Top             =   2220
      Width           =   1575
   End
   Begin VB.Label lblIncompleteText 
      Caption         =   "Incomplete Text"
      Height          =   195
      Left            =   60
      TabIndex        =   12
      Top             =   1620
      Width           =   1455
   End
   Begin VB.Label lblCompleteText 
      Caption         =   "Complete Text"
      Height          =   195
      Left            =   60
      TabIndex        =   11
      Top             =   1020
      Width           =   1335
   End
   Begin VB.Label lblBusinessArea 
      Caption         =   "Business Area"
      Height          =   255
      Left            =   60
      TabIndex        =   10
      Top             =   600
      Width           =   1395
   End
   Begin VB.Label lblGroupName 
      Caption         =   "Group Name"
      Height          =   195
      Left            =   60
      TabIndex        =   9
      Top             =   180
      Width           =   1155
   End
End
Attribute VB_Name = "frmEditBusinessGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form          : frmEditBusinessGroup
' Description   : Form which allows the adding/editing of Task Groups
'
' Change history
' Prog      Date        Description
' AA        09/03/01    Added Form
' DJP       20/06/01    SQL Server port - update no matter what.
' STB       08/07/2002  SYS4529 'ESC' now closes the form.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
' Private data
Private m_ReturnCode As MSGReturnCode
Private m_clsTaskLink As TaskLinkTable
Private m_clsBusinessGroup As BusinessGroupTable
Private m_clsTasks As TaskTable
Private m_bIsEdit As Boolean
Private m_colKeys As Collection
Private m_vGroupID As Variant

'ComboGroup Constants
Private Const COMBO_BUSINESS_AREA As String = "CTBusinessArea"

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SetIsEdit
' Description   :   Sets whether or not this form is in edit or add mode. Defaults to edit
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetIsEdit(Optional bIsEdit As Boolean = True)
    m_bIsEdit = bIsEdit
End Sub
Public Function GetReturnCode() As MSGReturnCode
    GetReturnCode = m_ReturnCode
End Function

Private Sub SetReturnCode(Optional enumReturn As MSGReturnCode = MSGSuccess)
    m_ReturnCode = enumReturn
End Sub

Private Sub cmdCancel_Click()
    Hide
End Sub

Private Sub cmdOK_Click()
    On Error GoTo Failed
    
    Dim bRet As Boolean
    
    bRet = DoOkProcessing
    
    If bRet Then
        SaveScreenData
        SaveChangeRequest
        SetReturnCode
        Hide
    End If
        
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError Err.DESCRIPTION
End Sub

Private Sub Form_Load()

    On Error GoTo Failed
        
    Set m_clsTasks = New TaskTable
    Set m_clsTaskLink = New TaskLinkTable
    Set m_clsBusinessGroup = New BusinessGroupTable
    
    m_vGroupID = ""
    
    PopulateScreenControls
    
   ' txtGroupID.Enabled = False
    
    If m_bIsEdit Then
        SetFormEditState
    End If
    
    PopulateAvailableItems swapTasks, enumStageAll
    
    Exit Sub
Failed:
    g_clsErrorHandling.DisplayError Err.DESCRIPTION
    g_clsFormProcessing.CancelForm Me
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : PopulateScreenControls
' Description   : Populate Combos
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PopulateScreenControls()
    On Error GoTo Failed
    
    'Populate Business Area Combo
    g_clsFormProcessing.PopulateCombo COMBO_BUSINESS_AREA, cboBusinessArea
    
    'Populate SwapList
    SetSwapListHeaders
    
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
    
End Sub

Private Sub SetSwapListHeaders()
    On Error GoTo Failed
    Dim headers As New Collection
    Dim lvHeaders As listViewAccess
    Dim colLine As Collection

    lvHeaders.nWidth = 80
    lvHeaders.sName = "Task Name"
    headers.Add lvHeaders

    swapTasks.SetFirstColumnHeaders headers
    swapTasks.SetSecondColumnHeaders headers

    swapTasks.SetFirstTitle "Task List"
    swapTasks.SetSecondTitle "Linked Tasks"
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub PopulateAvailableItems(swapList As MSGHorizontalSwapList, enumStageType As StageType)
    On Error GoTo Failed
    Dim colLine As Collection
    Dim clsSwapExtra As SwapExtra
    Dim rs As ADODB.Recordset
    Dim sTaskID As String
    Dim sTaskName As String
    Dim bExists As Boolean
    Dim clsTableAccess As TableAccess
    
    Set clsTableAccess = m_clsTasks
    
    Set rs = m_clsTaskLink.GetUnassignedTasks(CStr(m_vGroupID))
    clsTableAccess.SetRecordSet rs
    
    swapList.ClearFirst
    
    If clsTableAccess.RecordCount() > 0 Then
        
        clsTableAccess.MoveFirst
        
        Do While Not rs.EOF
            Set colLine = New Collection
            sTaskID = m_clsTasks.GetTaskID()
            sTaskName = m_clsTasks.GetTaskName

            bExists = g_clsFormProcessing.DoesSwapValueExist(swapList, sTaskName)
            
            If Not bExists And Len(sTaskName) > 0 Then
                Set clsSwapExtra = New SwapExtra
                clsSwapExtra.SetValueID sTaskID
                clsSwapExtra.SetAmount sTaskName
                
                colLine.Add sTaskName
                swapList.AddLineFirst colLine, clsSwapExtra
            
            End If
            
            rs.MoveNext
        Loop
        
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Public Sub SaveTasks()
    On Error GoTo Failed
    Dim vTmp As Variant
    Dim nSelectedCount As Integer
    Dim colValues As Collection
    Dim nThisItem As Integer
    Dim sValueID As String
    Dim sTaskName As String
    Dim clsSwapExtra As New SwapExtra
    
    If TableAccess(m_clsTaskLink).RecordCount > 0 Then
        TableAccess(m_clsTaskLink).DeleteAllRows
    End If
    
    nSelectedCount = swapTasks.GetSecondCount()

    For nThisItem = 1 To nSelectedCount
        Set colValues = swapTasks.GetLineSecond(nThisItem, clsSwapExtra)

        sValueID = clsSwapExtra.GetValueID
        sTaskName = clsSwapExtra.GetAmount
        
        If Len(sValueID) > 0 Then
            TableAccess(m_clsTaskLink).AddRow

            m_clsTaskLink.SetGroupID m_vGroupID
            m_clsTaskLink.SetTaskID sValueID
            m_clsTaskLink.SetTaskName sTaskName
        Else
            g_clsErrorHandling.RaiseError errGeneralError, "TaskID is empty"
        End If
    Next

    ' DJP SQL Server port - update no matter what.
    TableAccess(m_clsTaskLink).Update

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Function DoOkProcessing() As Boolean
    On Error GoTo Failed
    
    Dim bRet As Boolean
    
    bRet = g_clsFormProcessing.DoMandatoryProcessing(Me)
    
    DoOkProcessing = bRet
    
    Exit Function
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function

Private Sub SaveScreenData()
    On Error GoTo Failed
    Dim vVal As Variant
    Dim clsGUID As GuidAssist
    
    Set clsGUID = New GuidAssist
    
    If Not m_bIsEdit Then
        m_vGroupID = clsGUID.CreateGUID
        g_clsFormProcessing.CreateNewRecord m_clsBusinessGroup
        m_clsBusinessGroup.SetGroupID m_vGroupID
    End If
    
    g_clsFormProcessing.HandleComboExtra cboBusinessArea, vVal, GET_CONTROL_VALUE
    m_clsBusinessGroup.SetBusinessArea vVal
    
    m_clsBusinessGroup.SetGroupName txtGroupName.Text
    
    m_clsBusinessGroup.SetCompleteText txtCompleteText.Text
    
    m_clsBusinessGroup.SetInCompleteText txtIncompleteText.Text
    
    m_clsBusinessGroup.SetNotApplicableText txtNotApplicableText.Text
    
    m_clsBusinessGroup.SetNotStartedText txtNotStartedText.Text
    
    SaveTasks
    
    TableAccess(m_clsBusinessGroup).Update

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Public Sub SetKeys(colKeys As Collection)
    Set m_colKeys = colKeys
End Sub

Private Sub SetFormEditState()
    On Error GoTo Failed
    Dim clsTableAccess As TableAccess
    Dim rs As ADODB.Recordset
    Dim bRet As Boolean
    
    Set clsTableAccess = m_clsBusinessGroup
    
    clsTableAccess.SetKeyMatchValues m_colKeys
    
    Set rs = clsTableAccess.GetTableData(POPULATE_KEYS)
    bRet = True
    
    If Not rs Is Nothing Then
        If rs.RecordCount = 0 Then
            bRet = False
        End If
    Else
        bRet = False
    End If
    
    If Not bRet Then
        g_clsErrorHandling.RaiseError errRecordSetEmpty
    Else
        PopulateScreenFields
    End If
    
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

Private Sub PopulateScreenFields()
    On Error GoTo Failed
    Dim vVal As Variant
        
    
    ' DJP SQL Server port
    m_vGroupID = m_clsBusinessGroup.GetGroupID
    'Group Name
    txtGroupName.Text = m_clsBusinessGroup.GetGroupName
    
    'Business Area
    vVal = m_clsBusinessGroup.GetBusinessArea
    g_clsFormProcessing.HandleComboExtra cboBusinessArea, vVal, SET_CONTROL_VALUE
    
    'Complete Text
    txtCompleteText.Text = m_clsBusinessGroup.GetCompleteText
    
    'Incomplete Text
    txtIncompleteText.Text = m_clsBusinessGroup.GetInCompleteText
    
    'Not Applicable Text
    txtNotApplicableText.Text = m_clsBusinessGroup.GetNotApplicableText
    
    'Not Started Text
    txtNotStartedText.Text = m_clsBusinessGroup.GetNotStartedText
    
    'Task Swap List
    PopulateSelectedTasks
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub PopulateSelectedTasks()
    On Error GoTo Failed
    Dim bExists As Boolean
    Dim sTaskID As String
    Dim sTaskName As String
    Dim rs As ADODB.Recordset
    Dim colLine As Collection
    Dim clsSwapExtra As SwapExtra
    Dim clsTableAccess As TableAccess
    Dim colValues As Collection
    
    Set colValues = New Collection
    colValues.Add m_vGroupID
    Set clsTableAccess = m_clsTaskLink
    
    ' DJP SQL Server port. GetTasksForGroup not needed - the default GetTableData with keys willdo the job
    ' and format the GUID correctly.
    
    TableAccess(m_clsTaskLink).SetKeyMatchValues colValues
    TableAccess(m_clsTaskLink).GetTableData
    'm_clsTaskLink.GetTasksForGroup CStr(m_vGroupID)
    
    If clsTableAccess.RecordCount() > 0 Then
        Set rs = clsTableAccess.GetRecordSet()
    
        rs.MoveFirst
    
        Do While Not rs.EOF
            Set colLine = New Collection
            sTaskID = m_clsTaskLink.GetTaskID()
            sTaskName = m_clsTaskLink.GetTaskName
        
            bExists = g_clsFormProcessing.DoesSwapValueExist(swapTasks, sTaskName)
        
            If Not bExists Then
                Set clsSwapExtra = New SwapExtra
                clsSwapExtra.SetValueID sTaskID
                clsSwapExtra.SetAmount sTaskName
            
                colLine.Add sTaskName
                swapTasks.AddLineSecond colLine, clsSwapExtra
            End If
        
            rs.MoveNext
        Loop
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SaveChangeRequest
' Description   :   Common Function used to setup a promotion
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveChangeRequest()
    On Error GoTo Failed
    Dim colValues As Collection
    Dim clsTableAccess As TableAccess
    
    Set clsTableAccess = m_clsBusinessGroup
    
    Set colValues = New Collection
    colValues.Add m_vGroupID
    
    clsTableAccess.SetKeyMatchValues colValues
    
    g_clsHandleUpdates.SaveChangeRequest clsTableAccess, txtGroupName.Text
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
