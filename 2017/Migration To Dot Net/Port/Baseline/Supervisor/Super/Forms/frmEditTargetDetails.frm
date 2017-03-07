VERSION 5.00
Object = "{61B85B2D-AE68-4B98-B8E9-B759DBAF18E8}#1.1#0"; "MSGOCX.ocx"
Begin VB.Form frmEditTargetDetails 
   Caption         =   "Add/Edit Target Details"
   ClientHeight    =   2190
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4050
   Icon            =   "frmEditTargetDetails.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   4050
   StartUpPosition =   1  'CenterOwner
   Begin MSGOCX.MSGEditBox txtTarget 
      Height          =   315
      Index           =   0
      Left            =   1800
      TabIndex        =   0
      Top             =   120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      Mask            =   "##/##/####"
      Format          =   "c"
      TextType        =   1
      FontSize        =   8.25
      FontName        =   "MS Sans Serif"
      Mandatory       =   -1  'True
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
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   1380
      TabIndex        =   4
      Top             =   1740
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   180
      TabIndex        =   3
      Top             =   1740
      Width           =   1095
   End
   Begin MSGOCX.MSGEditBox txtTarget 
      Height          =   315
      Index           =   1
      Left            =   1800
      TabIndex        =   1
      Top             =   600
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      Mask            =   "##/##/####"
      Format          =   "c"
      TextType        =   1
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
   Begin MSGOCX.MSGEditBox txtTarget 
      Height          =   315
      Index           =   2
      Left            =   1800
      TabIndex        =   2
      Top             =   1080
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   556
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
      MaxLength       =   10
   End
   Begin VB.Label lbltarget 
      Caption         =   "Target"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   1140
      Width           =   1455
   End
   Begin VB.Label lbltarget 
      Caption         =   "Active To Date"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label lbltarget 
      Caption         =   "Active From Date"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   180
      Width           =   1455
   End
End
Attribute VB_Name = "frmEditTargetDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Form          :   frmEditIntReportDetails
' Description   :   Form which allows the user add and edit Intermediary Report details
'
' Change history
' Prog      Date        Description
' STB       15/11/01    SYS2550 SQL Server support.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

'A status indicator to the form's caller.
Private m_uReturnCode As MSGReturnCode

'The state of the form. True - edit mode, False - add mode.
Private m_bIsEdit As Boolean

'A collection of primary keys to identify the current record.
Private m_colKeys As Collection

'Underlying table object for the record.
Private m_clsTargetTable As IntermediaryTargetTable

'Control indexes.
Private Const ACTIVE_FROM         As Long = 0
Private Const ACTIVE_TO           As Long = 1
Private Const INTERMEDIARY_TARGET As Long = 2


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cmdOK_Click
' Description   : Validate and save the record, closing the form if everything is okay.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOK_Click()
    
    Dim bRet As Boolean
    
    On Error GoTo Failed

    'Validate data.
    bRet = DoOkProcessing

    If bRet Then
        'Save the control values to the underlying table.
        SaveScreenData
        
        'Indicate to the opener that everything has worked.
        SetReturnCode MSGSuccess
        
        'Hide the form and return control to the opener.
        Hide
    End If

    Exit Sub

Failed:
    g_clsErrorHandling.DisplayError Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : cmdCancel_Click
' Description   : Hide the form and return control to the caller. Also, undo the current
'                 change to the underlying record.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdCancel_Click()
    
    On Error GoTo Failed
        
    'If we're adding, we'll need to undo the adding of the record.
    If Not m_bIsEdit Then
        TableAccess(m_clsTargetTable).GetRecordSet.CancelBatch adAffectCurrent
    End If
    
    'Hide the form and return control to the opener.
    Hide
    
    Exit Sub

Failed:
    g_clsErrorHandling.DisplayError Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      :   SetIsEdit
' Description   :   Sets whether or not this form is in edit or add mode. Defaults to edit
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetIsEdit(Optional ByVal bIsEdit As Boolean = True)
    m_bIsEdit = bIsEdit
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : GetReturnCode
' Description   : Return the success code to the caller.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetReturnCode() As MSGReturnCode
    GetReturnCode = m_uReturnCode
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetReturnCode
' Description   : Sets the return code for the form for any calling method to check. Defaults
'                 to MSG_SUCCESS
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetReturnCode(Optional ByVal enumReturn As MSGReturnCode = MSGSuccess)
    m_uReturnCode = enumReturn
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetKeys
' Description   : Sets a collection of primary key fields at module-level.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetKeys(ByRef colKeys As Collection)
    Set m_colKeys = colKeys
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : Form_Load
' Description   : Setup the form according to its current state.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
    
    On Error GoTo Failed
    
    If m_bIsEdit Then
        SetEditState
    Else
        SetAddState
    End If

    Exit Sub
    
Failed:
    g_clsErrorHandling.DisplayError Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetTarget
' Description   : Associate the target table specified with the form.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetTarget(ByRef clsTargetTable As IntermediaryTargetTable)
    Set m_clsTargetTable = clsTargetTable
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetScreenData
' Description   : Populate the screen controls with data from the underlying table object.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetScreenData()
    
    On Error GoTo Failed

    'Active from Date.
    g_clsFormProcessing.HandleDate txtTarget(ACTIVE_FROM), m_clsTargetTable.GetActiveFrom, SET_CONTROL_VALUE

    'Active To Date.
    g_clsFormProcessing.HandleDate txtTarget(ACTIVE_TO), m_clsTargetTable.GetActiveTo, SET_CONTROL_VALUE
     
    'Target.
    txtTarget(INTERMEDIARY_TARGET).Text = m_clsTargetTable.GetTarget
          
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetEditState
' Description   : Populate the screen from the data.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetEditState()
    
    On Error GoTo Failed

    SetScreenData
       
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SetAddState
' Description   : Add a row to the underlying table (cancelling will rollback).
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetAddState()
    
    On Error GoTo Failed

    'Add a new row to the table object.
    TableAccess(m_clsTargetTable).AddRow
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : ValidateDates
' Description   : Validates the active from and to dates, ensuring the from < to date
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ValidateDates() As Boolean
    
    Dim dTo As Date
    Dim dFrom As Date
    Dim bRet As Boolean
    
    On Error GoTo Failed
    
    bRet = True
    
    'Get the from date
    g_clsFormProcessing.HandleDate txtTarget(ACTIVE_FROM), dFrom, GET_CONTROL_VALUE
    
    'Get the To Date
    If Len(txtTarget(ACTIVE_TO).ClipText) > 0 Then
        g_clsFormProcessing.HandleDate txtTarget(ACTIVE_TO), dTo, GET_CONTROL_VALUE
    Else
        dTo = 0
    End If
    
    If dFrom > dTo And dTo > 0 Then
        g_clsErrorHandling.DisplayError "Active To date must be greater than the Active From date"
        bRet = False
    End If
    
    ValidateDates = bRet
    
    Exit Function
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : DoOkProcessing
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function DoOkProcessing() As Boolean
    
    Dim bRet As Boolean
    
    On Error GoTo Failed
    
    bRet = g_clsFormProcessing.DoMandatoryProcessing(Me, True)
    
    If bRet Then
        bRet = ValidateDates
    End If
    
    DoOkProcessing = bRet
    
    Exit Function
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function      : SaveScreenData
' Description   :
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveScreenData()
    
    On Error GoTo Failed

    Dim vVal As Variant
    
    'Sequence Number
    vVal = m_clsTargetTable.GetNextSequenceNumber
    m_clsTargetTable.SetSequence vVal
    
    'Active from
    g_clsFormProcessing.HandleDate txtTarget(ACTIVE_FROM), vVal, GET_CONTROL_VALUE
    m_clsTargetTable.SetActiveFrom vVal
    
    'Active to
    g_clsFormProcessing.HandleDate txtTarget(ACTIVE_TO), vVal, GET_CONTROL_VALUE
    m_clsTargetTable.SetActiveTo vVal

    'Target
    m_clsTargetTable.SetTarget txtTarget(INTERMEDIARY_TARGET).Text
    
    'GUID
    m_clsTargetTable.SetIntermediaryGuid m_colKeys(INTERMEDIARY_KEY)
    
    Exit Sub
    
Failed:
    g_clsErrorHandling.RaiseError Err.Number, Err.DESCRIPTION
End Sub

