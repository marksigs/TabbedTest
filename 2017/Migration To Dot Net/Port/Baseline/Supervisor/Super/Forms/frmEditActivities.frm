VERSION 5.00
Object = "{61B85B2D-AE68-4B98-B8E9-B759DBAF18E8}#2.0#0"; "MSGOCX.ocx"
Begin VB.Form frmEditActivities 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add/Edit Activity"
   ClientHeight    =   6795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8730
   Icon            =   "frmEditActivities.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   8730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   60
      TabIndex        =   4
      Top             =   6300
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1260
      TabIndex        =   5
      Top             =   6300
      Width           =   1095
   End
   Begin VB.CommandButton cmdAnother 
      Caption         =   "&Another"
      Height          =   375
      Left            =   7440
      TabIndex        =   6
      Top             =   6300
      Width           =   1095
   End
   Begin MSGOCX.MSGHorizontalSwapList swapStages 
      Height          =   2055
      Left            =   120
      TabIndex        =   3
      Top             =   2040
      Width           =   8475
      _ExtentX        =   14949
      _ExtentY        =   3625
   End
   Begin MSGOCX.MSGTextMulti txtDescription 
      Height          =   1095
      Left            =   1680
      TabIndex        =   2
      Top             =   840
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   1931
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
   Begin MSGOCX.MSGEditBox txtActivity 
      Height          =   315
      Index           =   0
      Left            =   1680
      TabIndex        =   0
      Top             =   120
      Width           =   2115
      _ExtentX        =   3731
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
      MaxLength       =   20
   End
   Begin MSGOCX.MSGEditBox txtActivity 
      Height          =   315
      Index           =   1
      Left            =   1680
      TabIndex        =   1
      Top             =   480
      Width           =   4395
      _ExtentX        =   7752
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
   Begin MSGOCX.MSGHorizontalSwapList swapExceptionStages 
      Height          =   1995
      Left            =   120
      TabIndex        =   10
      Top             =   4200
      Width           =   8475
      _ExtentX        =   14949
      _ExtentY        =   3519
   End
   Begin VB.Label Label3 
      Caption         =   "Description"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   900
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Activity ID"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   180
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Activity Name"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   540
      Width           =   1335
   End
End
Attribute VB_Name = "frmEditActivities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form          :   frmEditACtivities
' Description   :   Form which allows the adding/editing of Activities for Task Management

' Change history
' Prog      Date        Description
' DJP       09/11/00    Created for Phase 2 Task Management
' STB       06/12/01    SYS1942 - Another button commits current transaction.
' STB       08/07/2002  SYS4529 BorderStyle set to 'Fixed Dialog'.
' DJP       27/02/2003  BM0354 Don't lock Stage table on Edit.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

' Private data
Private m_bIsEdit As Boolean
Private m_clsStages As StageTable
Private m_clsActivityTable As ActivityTable
Private m_sActivityID As String
Private m_ReturnCode As MSGReturnCode
Private m_colKeys As Collection

' Private constants
Private Const ACTIVITY_ID = 0
Private Const ACTIVITY_NAME = 1

' private enums
Private Enum ActivityOrdered
    enumSaveActivityOrder = 1
    enumDoNotSaveActivityOrder
End Enum
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SetIsEdit
' Description   :   Sets whether or not this form is in edit or add mode. Defaults to edit
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetIsEdit(Optional bIsEdit As Boolean = True)
    m_bIsEdit = bIsEdit
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   cmdAnother_Click
' Description   :   Called when the user clicks on the Another button. Need to perform the same
'                   processing as if the user has pressed ok. Do validation, save screen data,
'                   clear the screen of existing values and put the focus on the TASK ID field
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdAnother_Click()
    On Error GoTo Failed
        
    ' Perform ok processing. Will raise an error if it fails
    DoOKProcessing
    
    'If the record was saved, commit the transaction and begin another.
    g_clsDataAccess.CommitTrans
    g_clsDataAccess.BeginTrans
    
    ' Clear all fields, set the default focus
    g_clsFormProcessing.ClearScreenFields Me
    
    Initialise
    'SetAddState
    
    txtActivity(ACTIVITY_ID).SetFocus
    
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
    Set m_clsStages = New StageTable
    Set m_clsActivityTable = New ActivityTable
    m_ReturnCode = MSGFailure
    
    InitialiseSwapList
    
    ' Edit mode?
    If m_bIsEdit Then
        SetEditState
    Else
        SetAddState
    End If
    
    PopulateStages
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub InitialiseSwapList()
    On Error GoTo Failed
    
    ' Exception stages
    swapExceptionStages.ClearFirst
    swapExceptionStages.ClearSecond
    swapExceptionStages.SetFirstSorted
    swapExceptionStages.SetSecondSorted
    swapStages.AllowReorder False
    
    ' Non-exception stages
    swapStages.ClearFirst
    swapStages.ClearSecond
    
    swapStages.SetFirstSorted True
    swapStages.SetSecondSorted False
    swapStages.AllowReorder True
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   PopulateStages
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PopulateStages()
    On Error GoTo Failed
    
    SetStageHeaders
    SetExceptionStageHeaders
    
    PopulateAvailableItems swapStages, enumDefaultStage
    PopulateAvailableItems swapExceptionStages, enumExceptionStage
    
    Exit Sub
Failed:
    
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub SetStageHeaders()
    On Error GoTo Failed
    Dim headers As New Collection
    Dim lvHeaders As listViewAccess
    Dim colLine As Collection

    lvHeaders.nWidth = 100
    lvHeaders.sName = "Stage Name"
    headers.Add lvHeaders

    swapStages.SetFirstColumnHeaders headers
    swapStages.SetSecondColumnHeaders headers

    swapStages.SetFirstTitle "Select Stage(s)"
    swapStages.SetSecondTitle "Selected Stage(s)"
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub SetExceptionStageHeaders()
    On Error GoTo Failed
    Dim headers As New Collection
    Dim lvHeaders As listViewAccess
    Dim colLine As Collection

    lvHeaders.nWidth = 100
    lvHeaders.sName = "Exception Stage Name"
    headers.Add lvHeaders

    swapExceptionStages.SetFirstColumnHeaders headers
    swapExceptionStages.SetSecondColumnHeaders headers

    swapExceptionStages.SetFirstTitle "Select Exception Stage(s)"
    swapExceptionStages.SetSecondTitle "Selected Exception Stage(s)"
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub PopulateSelectedItems()
    On Error GoTo Failed
    Dim nException As Long
    Dim rsStage  As ADODB.Recordset
    Dim clsSwapExtra As SwapExtra
    Dim colLine As Collection
    Dim sStageID As String
    Dim sStageName As String
    
    Set colLine = New Collection
    
    ' DJP BM0354 Return exception and non exception stages in one go and check later
    m_clsStages.GetStagesForActivity m_sActivityID, enumStageAll
    
    If TableAccess(m_clsStages).RecordCount() > 0 Then
        Set rsStage = TableAccess(m_clsStages).GetRecordSet()
        
        rsStage.MoveFirst
        
        ' We need to set the activityID to NULL for all of these incase they are removed
        Do While Not rsStage.EOF
            Set colLine = New Collection
            Set clsSwapExtra = New SwapExtra
            
            sStageID = m_clsStages.GetStageID()
            sStageName = m_clsStages.GetStageName()
            
            Set clsSwapExtra = New SwapExtra
            clsSwapExtra.SetValueID sStageID
            
            colLine.Add sStageName
        
            nException = m_clsStages.GetException
            
            ' DJP BM0354 Add to the right swaplist depending on whether it's an exception stage or not.
            If nException = 0 Then
                swapStages.AddLineSecond colLine, clsSwapExtra
            ElseIf nException = 1 Then
                swapExceptionStages.AddLineSecond colLine, clsSwapExtra
            End If
            
            ' Set to NULL
            m_clsStages.SetActivityID Null
            
            rsStage.MoveNext
        Loop
        
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub PopulateAvailableItems(SwapList As MSGHorizontalSwapList, enumStageType As StageType)
    On Error GoTo Failed
    Dim clsStages As StageTable
    Dim colLine As Collection
    Dim clsSwapExtra As SwapExtra
    Dim rsStage As ADODB.Recordset
    Dim sStageID As String
    Dim sStageName As String
    Dim bExists As Boolean
    
    Set clsStages = New StageTable

    clsStages.GetStages , , True, enumStageType
        
    If TableAccess(clsStages).RecordCount() > 0 Then
        Set rsStage = TableAccess(clsStages).GetRecordSet()
        
        rsStage.MoveFirst
        
        Do While Not rsStage.EOF
            Set colLine = New Collection
            sStageID = clsStages.GetStageID()
            sStageName = clsStages.GetStageName()
            
            bExists = g_clsFormProcessing.DoesSwapValueExist(SwapList, sStageName)
            
            If Not bExists Then
                Set clsSwapExtra = New SwapExtra
                clsSwapExtra.SetValueID sStageID
                
                colLine.Add sStageName
                SwapList.AddLineFirst colLine, clsSwapExtra
            
            End If
            
            rsStage.MoveNext
        Loop
        
        TableAccess(clsStages).CloseRecordSet
        
    End If
    
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
    
    g_clsFormProcessing.DoMandatoryProcessing Me
    
    If Not m_bIsEdit Then
        Dim sActivityID As String
        Dim bExists As Boolean
        Dim colMatchValues As Collection
               
        Set colMatchValues = New Collection
        
        sActivityID = txtActivity(ACTIVITY_ID).Text
        colMatchValues.Add sActivityID
        
        bExists = TableAccess(m_clsActivityTable).DoesRecordExist(colMatchValues)
    
        If bExists Then
            txtActivity(ACTIVITY_ID).SetFocus
            g_clsErrorHandling.RaiseError errGeneralError, "Activity already exists. Please enter a unique Activity ID"
        End If
    
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
    colValues.Add Me.txtActivity(ACTIVITY_ID).Text

    TableAccess(m_clsActivityTable).SetKeyMatchValues colValues

    ' Save this key
    g_clsHandleUpdates.SaveChangeRequest m_clsActivityTable
    
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
    Dim sActivityID As String
    Dim vVal As Variant
    
    m_sActivityID = txtActivity(ACTIVITY_ID).Text
    
    ' Activity ID
    m_clsActivityTable.SetActivityID m_sActivityID
    
    ' Activity Name
    m_clsActivityTable.SetActivityName txtActivity(ACTIVITY_NAME).Text
    
    ' Description
    m_clsActivityTable.SetActivityDescription txtDescription.Text
    
    m_clsActivityTable.SetDeleteFlag
    
    ' DJP BM0354 Before updating the stages, clear the Activity GUID's for each stage. This was done
    ' on screen entry, but to not lock the stage table, update it here first.
    If m_bIsEdit Then
        TableAccess(m_clsStages).Update
        TableAccess(m_clsStages).CloseRecordSet
    End If
    
    SaveStagesForActivity swapStages, enumSaveActivityOrder
    SaveStagesForActivity swapExceptionStages, enumDoNotSaveActivityOrder
    
    TableAccess(m_clsActivityTable).Update
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub SaveStagesForActivity(SwapList As MSGHorizontalSwapList, enumActivityOrder As ActivityOrdered)
    On Error GoTo Failed
    Dim sStageID As String
    Dim clsStage As StageTable
    Dim nThisItem As Integer
    Dim nSelectedCount As Long
    Dim colLine As Collection
    Dim colStages As Collection
    Dim clsSwapExtra As SwapExtra
    
    Set colStages = New Collection
    nSelectedCount = SwapList.GetSecondCount()
    
    For nThisItem = 1 To nSelectedCount
        Set colLine = SwapList.GetLineSecond(nThisItem, clsSwapExtra)

        If Not clsSwapExtra Is Nothing Then
            sStageID = clsSwapExtra.GetValueID()
            colStages.Add sStageID
        End If
    Next
        
    If colStages.Count > 0 Then
        Set clsStage = New StageTable
        
        clsStage.GetStages colStages
            
        If TableAccess(clsStage).RecordCount() > 0 Then
            Dim rsStage As ADODB.Recordset
            
            Set rsStage = TableAccess(clsStage).GetRecordSet()
             
            rsStage.MoveFirst
            
            Do While Not rsStage.EOF
                clsStage.SetActivityID m_sActivityID
                
                If enumActivityOrder = enumSaveActivityOrder Then
                    ' Set the Stage Sequence Number
                    SetOrderedStage colStages, clsStage
                Else
                    clsStage.SetStageSequenceNumber Null
                End If
                
                rsStage.MoveNext
                nThisItem = nThisItem + 1
            Loop
        End If
        
        TableAccess(clsStage).Update
        TableAccess(clsStage).CloseRecordSet
        
    End If
    
    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub SetOrderedStage(colStages As Collection, clsStage As StageTable)
    On Error GoTo Failed
    Dim bFound As Boolean
    Dim nTotal As Long
    Dim nThisItem As Long
    Dim sStageID As String
    Dim sOrderedStage As String
    
    bFound = False
    nThisItem = 1
    
    nTotal = colStages.Count
    sStageID = clsStage.GetStageID()
        
    Do While nThisItem <= nTotal And Not bFound
        sOrderedStage = colStages(nThisItem)
        
        If sStageID = sOrderedStage Then
            clsStage.SetStageSequenceNumber nThisItem
            bFound = True
        End If
        
        nThisItem = nThisItem + 1
    Loop
    
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
    
    Dim vVal As Variant
    
    ' Activity ID
    m_sActivityID = m_clsActivityTable.GetActivityID()
    txtActivity(ACTIVITY_ID).Text = m_sActivityID
    
    ' Activity Name
    txtActivity(ACTIVITY_NAME).Text = m_clsActivityTable.GetActivityName()
    
    ' Description
    txtDescription.Text = m_clsActivityTable.GetActivityDescription()
    
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
    
    g_clsFormProcessing.CreateNewRecord m_clsActivityTable
    
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
    
    TableAccess(m_clsActivityTable).SetKeyMatchValues m_colKeys
    
    ' Get the data from the database
    TableAccess(m_clsActivityTable).GetTableData
    
    ' Validate we have the record
    If TableAccess(m_clsActivityTable).RecordCount() = 0 Then
        g_clsErrorHandling.RaiseError errGeneralError, "Edit Tasks - Unable to locate task"
    End If

    ' If we get here, we have the data we need
    PopulateScreenFields
    
    ' Populate the swaplist controls
    PopulateSelectedItems
    
    ' Any other field enabling/disabling
    cmdAnother.Enabled = False
    txtActivity(ACTIVITY_ID).Enabled = False
    
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
    
    sFunctionName = "SetKeys"
    
    If Not colKeys Is Nothing Then
        If colKeys.Count > 0 Then
            Set m_colKeys = colKeys
        Else
            g_clsErrorHandling.RaiseError errKeysEmpty, "EditTasks, " & sFunctionName
        End If
    Else
        g_clsErrorHandling.RaiseError errKeysEmpty, "EditTasks, " & sFunctionName
    End If

    Exit Sub
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub
Private Sub Form_Initialize()
    m_bIsEdit = False
End Sub


